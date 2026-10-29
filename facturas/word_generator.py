import jinja2
from PIL import Image
import os
import io
import re
import copy
import zipfile
import datetime
import uuid
from decimal import Decimal
from django.conf import settings
from django.utils import timezone
import docx
from docx.shared import Mm, Pt
from docx.enum.text import WD_COLOR_INDEX
from docxtpl import DocxTemplate, InlineImage
import qrcode
import barcode
from barcode.writer import ImageWriter

PLANTILLA_ALIASES = {
    "EXITO": ["EXITO", "ALMACENES EXITO", "ALMACENES EXITO SA", "ALMACENES EXITO S.A", "EXITO WOW"],
    "FRISBY": ["FRISBY", "FRISBY SA", "FRISBY S.A", "FRISBY S.A. BIC", "FRISBY BIC"],
    "HOMECENTER": ["HOMECENTER", "SODIMAC", "SODIMAC COLOMBIA", "SODIMAC COLOMBIA SA"],
    "JUMBO": ["JUMBO", "CENCOSUD", "CENCOSUD COLOMBIA", "CENCOSUD COLOMBIA S.A", "JUMBO SANTA FE"],
    "OLIMPICA": ["OLIMPICA", "OLIMPICA SA", "OLIMPICA S.A"]
}

def normalizar_texto(texto):
    if not texto:
        return ""
    import unicodedata
    texto_norm = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8')
    return texto_norm.upper().strip()

def resolver_plantilla_factura(factura):
    """
    Determina que plantilla usar:
    1. Si el lote tiene factura_guia adjunta valida en disco, usa esa.
    2. Si la factura tiene explicitamente el campo 'plantilla' (JUMBO, EXITO, FRISBY, etc.), usa esa.
    3. Si coincide con alguno de los emisores registrados en plantillas_base, usa esa.
    4. Fallback seguro a cualquier plantilla base disponible en plantillas_base/.
    """
    if factura.lote_generacion and getattr(factura.lote_generacion, 'factura_guia', None):
        try:
            if factura.lote_generacion.factura_guia.name and os.path.exists(factura.lote_generacion.factura_guia.path):
                return factura.lote_generacion.factura_guia.path
        except Exception:
            pass

    plantillas_dir = os.path.join(settings.BASE_DIR, 'plantillas_base')

    # Chequear campo plantilla explicito con validación inteligente frente al emisor
    plantilla_val = getattr(factura, 'plantilla', None)
    emisor_raw = (getattr(factura, 'emisor_nombre', '') or '').upper()
    nit_raw = (getattr(factura, 'emisor_nit', '') or '').replace('.', '').replace('-', '').strip()
    if plantilla_val == 'JUMBO' and ('EXITO' in emisor_raw or 'ÉXITO' in emisor_raw or '890900608' in nit_raw):
        plantilla_val = 'EXITO'
    elif plantilla_val == 'EXITO' and ('CENCOSUD' in emisor_raw or 'JUMBO' in emisor_raw or '900155107' in nit_raw):
        plantilla_val = 'JUMBO'

    if plantilla_val:
        p = os.path.join(plantillas_dir, f"{str(plantilla_val).upper().strip()}.docx")
        if os.path.exists(p):
            return p

    emisor_norm = normalizar_texto(factura.emisor_nombre)

    # Busqueda por alias
    for nombre_plantilla, alias_list in PLANTILLA_ALIASES.items():
        for alias in alias_list:
            if normalizar_texto(alias) in emisor_norm:
                p = os.path.join(plantillas_dir, f"{nombre_plantilla}.docx")
                if os.path.exists(p):
                    return p

    # Fallback por defecto garantizado
    for fb in ['JUMBO.docx', 'EXITO.docx', 'FRISBY.docx', 'HOMECENTER.docx', 'OLIMPICA.docx']:
        fb_path = os.path.join(plantillas_dir, fb)
        if os.path.exists(fb_path):
            return fb_path

    # Ultimo recurso: primer archivo docx en plantillas_base
    try:
        if os.path.exists(plantillas_dir):
            for f in os.listdir(plantillas_dir):
                if f.lower().endswith('.docx'):
                    return os.path.join(plantillas_dir, f)
    except Exception:
        pass

    return None


def limpiar_monto_colombiano(val_str):
    if not val_str:
        return Decimal('0')
    s = str(val_str).strip()
    s = re.sub(r'[\$A-Za-z\s]', '', s)
    if '.' in s and ',' not in s:
        partes = s.split('.')
        if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
            s = s.replace('.', '')
        elif len(partes) == 2 and len(partes[1]) in [1, 2]:
            pass
        else:
            s = s.replace('.', '')
    elif ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s and '.' not in s:
        partes = s.split(',')
        if len(partes) == 2 and len(partes[1]) == 3:
            s = s.replace(',', '')
        elif len(partes) == 2 and len(partes[1]) in [1, 2]:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
            
    match = re.search(r'[\d\.]+', s)
    if match:
        try:
            return Decimal(match.group(0))
        except Exception:
            return Decimal('0')
    return Decimal('0')

def parsear_productos_texto(texto):
    """
    Parsea texto plano con productos y devuelve lista de dicts:
    [{'codigo': str, 'nombre': str, 'cantidad': int, 'valor_unitario': Decimal, 'total_linea': Decimal}, ...]
    Soporta formato Excel (tabulado), pipes (|), comas, formato natural (1 x 3.365.000 Nombre)
    y tirillas oficiales copiadas directamente.
    """
    if not texto or not texto.strip():
        return []

    lineas = [l.strip() for l in texto.strip().splitlines() if l.strip()]
    productos = []
    
    i = 0
    while i < len(lineas):
        linea = lineas[i]
        
        # Ignorar encabezados comunes si vienen en el copiado
        if any(h in linea.upper() for h in ['PLU', 'DETALLE', 'PRECIO', 'TOTAL ITEM', 'SUBTOTAL', 'VALOR TOTAL', 'FORMA PAGO', 'DISCRIMINACION', 'ITEMS COMPRADOS', 'RESUMEN DE IVA']):
            i += 1
            continue
            
        m_exito_header = re.match(r'^\d+\s+(\d+)/u\s+[xX]\s*([\$\d\.,]+)', linea)
        if m_exito_header:
            cantidad = int(m_exito_header.group(1))
            precio = limpiar_monto_colombiano(m_exito_header.group(2))
            desc_lines = []
            i += 1
            while i < len(lineas):
                sig = lineas[i]
                if re.match(r'^\d+\s+(\d+)/u\s+[xX]', sig) or any(h in sig.upper() for h in ['TOTAL ITEM', 'SUBTOTAL']):
                    break
                desc_lines.append(sig)
                i += 1
                if any(sig.endswith(suf) for suf in ['A', 'B', 'N']) and any(c.isdigit() for c in sig):
                    break
            
            desc_full = " ".join(desc_lines)
            m_plu = re.match(r'^(\d{5,8})\s+(.*)$', desc_full)
            if m_plu:
                codigo = m_plu.group(1)
                nombre_raw = m_plu.group(2)
            else:
                codigo = f"{random.randint(3100000, 3999999)}"
                nombre_raw = desc_full
            nombre = re.sub(r'[\$\d\.,]+[A-Za-z]?\s*$', '', nombre_raw).strip()
            if not nombre:
                nombre = nombre_raw.strip()
            total_linea = precio * cantidad
            productos.append({
                'codigo': codigo,
                'nombre': nombre,
                'cantidad': cantidad,
                'valor_unitario': precio,
                'total_linea': total_linea
            })
            continue

        delimitador = None
        for sep in ['\t', '|', ';']:
            if sep in linea:
                delimitador = sep
                break
        
        if delimitador:
            partes = [p.strip() for p in linea.split(delimitador) if p.strip()]
            if len(partes) >= 4:
                codigo = partes[0]
                nombre = partes[1]
                try:
                    cantidad = int(re.sub(r'\D', '', partes[2]) or '1')
                except Exception:
                    cantidad = 1
                precio = limpiar_monto_colombiano(partes[3])
            elif len(partes) == 3:
                if re.match(r'^\d+$', partes[1]) and len(partes[1]) <= 4:
                    codigo = f"{random.randint(3100000, 3999999)}"
                    nombre = partes[0]
                    cantidad = int(partes[1])
                    precio = limpiar_monto_colombiano(partes[2])
                elif re.match(r'^\d{5,8}$', partes[0]):
                    codigo = partes[0]
                    nombre = partes[1]
                    cantidad = 1
                    precio = limpiar_monto_colombiano(partes[2])
                else:
                    codigo = f"{random.randint(3100000, 3999999)}"
                    nombre = partes[0]
                    try:
                        cantidad = int(re.sub(r'\D', '', partes[1]) or '1')
                    except Exception:
                        cantidad = 1
                    precio = limpiar_monto_colombiano(partes[2])
            elif len(partes) == 2:
                codigo = f"{random.randint(3100000, 3999999)}"
                nombre = partes[0]
                cantidad = 1
                precio = limpiar_monto_colombiano(partes[1])
            else:
                i += 1
                continue
        else:
            m_nat = re.match(r'^(\d+)\s*[xX]\s*([\$\d\.,]+)\s+(.+)$', linea)
            if m_nat:
                cantidad = int(m_nat.group(1))
                precio = limpiar_monto_colombiano(m_nat.group(2))
                nombre = m_nat.group(3).strip()
                codigo = f"{random.randint(3100000, 3999999)}"
            else:
                m_coma = [p.strip() for p in linea.split(',') if p.strip()]
                if len(m_coma) >= 3:
                    codigo = f"{random.randint(3100000, 3999999)}"
                    nombre = m_coma[0]
                    try:
                        cantidad = int(re.sub(r'\D', '', m_coma[1]) or '1')
                    except Exception:
                        cantidad = 1
                    precio = limpiar_monto_colombiano(m_coma[2])
                elif len(m_coma) == 2:
                    codigo = f"{random.randint(3100000, 3999999)}"
                    nombre = m_coma[0]
                    cantidad = 1
                    precio = limpiar_monto_colombiano(m_coma[1])
                else:
                    m_fin = re.search(r'^(.*?)\s+([\$\d\.,]{3,})$', linea)
                    if m_fin:
                        codigo = f"{random.randint(3100000, 3999999)}"
                        nombre = m_fin.group(1).strip()
                        cantidad = 1
                        precio = limpiar_monto_colombiano(m_fin.group(2))
                    else:
                        i += 1
                        continue

        if nombre:
            if cantidad <= 0:
                cantidad = 1
            total_linea = precio * cantidad
            productos.append({
                'codigo': str(codigo).strip(),
                'nombre': str(nombre).strip(),
                'cantidad': cantidad,
                'valor_unitario': precio,
                'total_linea': total_linea
            })
        i += 1

    return productos

def formatear_lineas_producto_exito(idx, codigo, nombre, cantidad, valor_unitario, total_linea):
    lineas = []
    p_unit_str = fmt_pesos(valor_unitario, separador_miles=".")
    lineas.append(f"{idx} {cantidad}/u x {p_unit_str} V.Ahorro 0")
    
    precio_str = f"{fmt_pesos(total_linea, separador_miles='.')}A"
    texto_completo = f"{codigo} {nombre}".strip()
    palabras = texto_completo.split()
    
    lineas_desc = []
    linea_actual = ""
    max_ancho = 28
    
    for p in palabras:
        if not linea_actual:
            linea_actual = p
        elif len(linea_actual) + 1 + len(p) <= max_ancho:
            linea_actual += " " + p
        else:
            lineas_desc.append(linea_actual)
            linea_actual = p
    if linea_actual:
        lineas_desc.append(linea_actual)
        
    if not lineas_desc:
        lineas_desc = [texto_completo]
        
    ancho_total = 40
    ultima = lineas_desc[-1]
    
    if len(ultima) + 1 + len(precio_str) > ancho_total:
        lineas.extend(lineas_desc)
        espacios = ancho_total - len(precio_str)
        lineas.append(" " * max(espacios, 1) + precio_str)
    else:
        espacios = ancho_total - len(ultima) - len(precio_str)
        lineas_desc[-1] = ultima + (" " * max(espacios, 1)) + precio_str
        lineas.extend(lineas_desc)
        
    return lineas

def fmt_pesos(val, separador_miles="."):
    if val is None:
        return "0"
    try:
        val_int = int(round(Decimal(str(val))))
        s = f"{val_int:,}"
        if separador_miles == ".":
            s = s.replace(",", ".")
        return s
    except Exception:
        return str(val)

def generar_qr_stream(url):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    stream = io.BytesIO()
    img.save(stream, format="PNG")
    stream.seek(0)
    return stream

def generar_barcode_stream(code_text):
    clean_code = re.sub(r'[^A-Za-z0-9]', '', str(code_text))
    if not clean_code:
        clean_code = '09589'
    code128_cls = barcode.get_barcode_class('code128')
    bc = code128_cls(clean_code, writer=ImageWriter())
    stream = io.BytesIO()
    bc.write(stream, options={
        'write_text': False,
        'module_width': 0.28,
        'module_height': 12.0,
        'quiet_zone': 1.0,
    })
    stream.seek(0)
    return stream

def limpiar_resaltados(doc):
    """Elimina el fondo amarillo de todos los elementos XML (runs, oMath, pPr) para impresión profesional."""
    try:
        for elem in doc._element.xpath('.//w:highlight'):
            elem.getparent().remove(elem)
    except Exception:
        pass
    for p in doc.paragraphs:
        for r in p.runs:
            if r.font.highlight_color is not None:
                r.font.highlight_color = WD_COLOR_INDEX.AUTO

def rellenar_plantilla_jumbo(doc_path, factura):
    """
    Rellena la plantilla oficial de JUMBO respetando 100% formato, márgenes de 80mm y fuentes.
    Inyecta dinámicamente:
    - QR Code apuntando al dominio en Render
    - Código de barras Code128 de REM
    - CUFE, Folio, Serie, Fecha, Hora, ST, RG, TR, REM
    - Elimina cualquier resaltado amarillo
    """
    doc = docx.Document(doc_path)
    base_url = os.environ.get('APP_DOMAIN') or os.environ.get('RENDER_EXTERNAL_URL') or 'https://catalogo-vpfe1.onrender.com'
    base_url = base_url.rstrip('/')

    cufe = factura.cufe or str(uuid.uuid4()).replace('-', '') + str(uuid.uuid4()).replace('-', '')
    serie = factura.serie or "AU06"
    folio = factura.folio or "3425557"
    fecha = factura.fecha_emision or timezone.localdate()
    hora = factura.hora_emision or timezone.localtime(timezone.now()).time()

    fecha_corta = fecha.strftime('%d/ %m/ %y')
    hora_corta = hora.strftime('%H : %M')

    st = getattr(factura, 'terminal', '1034') or '1034'
    rg = getattr(factura, 'registro', '05') or '05'
    tr = getattr(factura, 'transaccion', '9983') or '9983'
    rem = getattr(factura, 'rem', f"REM{folio[-5:] if len(folio)>=5 else '09596'}") or 'REM09596'
    if not rem.startswith('REM'):
        rem = f"REM{rem}"

    # Reemplazo de imágenes en partes DOCX
    qr_stream = generar_qr_stream(f"{base_url}/?cufe={cufe}")
    qr_bytes = qr_stream.getvalue()
    if 'rId8' in doc.part.rels:
        doc.part.rels['rId8'].target_part._blob = qr_bytes

    # Se preserva exactamente el código de barras original de la plantilla de Jumbo
    pass

    # Reemplazo quirúrgico en párrafos y runs exactos preservando estilos y tipografía
    if len(doc.paragraphs) > 42:
        p42 = doc.paragraphs[42]
        if p42.runs:
            p42.runs[0].text = cufe
            for r in p42.runs[1:]:
                r.text = ""
        else:
            p42.text = cufe

    if len(doc.paragraphs) > 44:
        p44 = doc.paragraphs[44]
        if len(p44.runs) >= 3:
            p44.runs[1].text = f"{serie}- {folio}"
        else:
            p44.text = f"  FACTURADOR  :  {serie}- {folio}   - CUFE "

    if len(doc.paragraphs) > 45:
        p45 = doc.paragraphs[45]
        if p45.runs:
            p45.runs[0].text = cufe
            for r in p45.runs[1:]:
                r.text = ""
        else:
            p45.text = cufe

    if len(doc.paragraphs) > 52:
        p52 = doc.paragraphs[52]
        if len(p52.runs) >= 2:
            p52.runs[1].text = f"{serie}  {folio}"
        else:
            p52.text = f"   FAC.    ELECTRONICA   DE  VENTA  {serie}  {folio}"

    if len(doc.paragraphs) > 63:
        p63 = doc.paragraphs[63]
        if len(p63.runs) >= 13:
            p63.runs[0].text = f"{hora.hour:02d}"
            p63.runs[1].text = " : "
            p63.runs[2].text = f"{hora.minute:02d}"
            p63.runs[3].text = "      "
            p63.runs[4].text = f"{fecha.day:02d}"
            p63.runs[5].text = f"/ {fecha.month:02d}/ {fecha.strftime('%y')}"
            p63.runs[12].text = str(tr)
        else:
            p63.text = f"{hora_corta}      {fecha_corta}      ST {st} RG#{rg}       TR#{tr}"

    if len(doc.paragraphs) > 64:
        p64 = doc.paragraphs[64]
        if len(p64.runs) >= 2:
            p64.runs[0].text = "                "
            p64.runs[1].text = rem
            for r in p64.runs[2:]:
                r.text = ""
        else:
            p64.text = f"                {rem}"

    # Inyección dinámica de productos en plantilla JUMBO
    productos = factura.productos.all().order_by('id')
    if productos.exists():
        idx_inicio = None
        idx_fin = None
        for p_idx, p in enumerate(doc.paragraphs):
            if '=================================================' in p.text and idx_inicio is None:
                idx_inicio = p_idx
            elif '--------------------------------------------------------------------------------' in p.text and idx_inicio is not None and p_idx > idx_inicio:
                idx_fin = p_idx
                break

        if idx_inicio is not None and idx_fin is not None:
            target_p = doc.paragraphs[idx_fin]
            parrafos_a_borrar = [doc.paragraphs[k] for k in range(idx_inicio + 1, idx_fin)]
            for p in parrafos_a_borrar:
                p._p.getparent().remove(p._p)

            for prod in productos:
                p_unit_str = fmt_pesos(prod.valor_unitario)
                p_tot_str = fmt_pesos(prod.total_linea)
                cod = prod.codigo or "3669368"
                l1 = f"{prod.cantidad}  X {p_unit_str}"
                l2 = f"{cod} {prod.nombre[:45]}      ${p_tot_str}"
                new_p1 = target_p.insert_paragraph_before(l1)
                new_p1.style = doc.styles['Normal']
                new_p2 = target_p.insert_paragraph_before(l2)
                new_p2.style = doc.styles['Normal']
            new_p_space = target_p.insert_paragraph_before("      ")
            new_p_space.style = doc.styles['Normal']

        total_fmt = fmt_pesos(factura.total)
        subtotal_fmt = fmt_pesos(factura.subtotal or factura.total)
        iva_fmt = fmt_pesos(factura.iva)
        cambio_fmt = fmt_pesos(Decimal('50000.00') - factura.total) if factura.total < Decimal('50000.00') else '0'

        for p in doc.paragraphs:
            if 'ITEMS   COMPRADOS' in p.text:
                p.text = f"            ITEMS   COMPRADOS  : {productos.count()}"
            elif 'SUBTOTAL/TOTAL' in p.text:
                p.text = f"                   SUBTOTAL/TOTAL                                  {total_fmt}"
            elif p.text.startswith('EFECTIVO'):
                p.text = f"EFECTIVO                                                                             {total_fmt}"
            elif p.text.startswith('CAMBIO'):
                p.text = f"CAMBIO                                                                                 {cambio_fmt}"
            elif p.text.startswith('N=19%'):
                p.text = f"N=19%                                  {subtotal_fmt}                       {iva_fmt}"
            elif p.text.startswith(' TOTAL'):
                p.text = f" TOTAL                              {total_fmt}                           {iva_fmt}"

    limpiar_resaltados(doc)
    return doc

def rellenar_plantilla_exito(doc_path, factura):
    """
    Rellena la plantilla oficial de ÉXITO respetando 100% tipografía Courier New, márgenes, espaciado y alineación.
    Inyecta dinámicamente:
    - QR Code apuntando al dominio en Render
    - Código de barras Code128 del número de factura
    - CUFE, Número, Fecha, Hora, TR, RANG. AUT, MAPA, NUTRX
    - Elimina cualquier resaltado amarillo
    """
    doc = docx.Document(doc_path)
    base_url = os.environ.get('APP_DOMAIN') or os.environ.get('RENDER_EXTERNAL_URL') or 'https://catalogo-vpfe1.onrender.com'
    base_url = base_url.rstrip('/')

    cufe = factura.cufe or str(uuid.uuid4()).replace('-', '') + str(uuid.uuid4()).replace('-', '')
    serie = factura.serie or "WJ53"
    folio = factura.folio or "44660"
    numero_factura = f"{serie}{folio}"
    fecha = factura.fecha_emision or timezone.localdate()
    hora = factura.hora_emision or timezone.localtime(timezone.now()).time()

    fecha_iso = fecha.strftime('%Y-%m-%d')
    hora_iso = hora.strftime('%H:%M:%S')
    fecha_hora_str = f"{fecha_iso} {hora_iso}"
    fecha_compact = fecha.strftime('%Y%m%d')

    tr = getattr(factura, 'transaccion', '0197') or '0197'
    try:
        tr_int = int(re.sub(r'\D', '', str(tr)) or '197')
    except Exception:
        tr_int = 197
    tr_formatted = f"{tr_int:04d}"

    # Reemplazo de imágenes en partes DOCX (rId6 = QR, rId7 = Barcode)
    qr_stream = generar_qr_stream(f"{base_url}/?cufe={cufe}")
    qr_bytes = qr_stream.getvalue()
    if 'rId6' in doc.part.rels:
        doc.part.rels['rId6'].target_part._blob = qr_bytes

    # Se preserva exactamente el código de barras original de la plantilla de Éxito según requerimiento
    pass

    # Reemplazo en párrafos buscando por contenido exacto para máxima fidelidad con el espaciado
    for idx, p in enumerate(doc.paragraphs):
        p_text = p.text.strip()
        if not p_text:
            continue

        if p_text.startswith('NUMERO:'):
            if len(p.runs) >= 2:
                p.runs[1].text = numero_factura
            else:
                p.text = f"NUMERO: {numero_factura}"
        elif re.search(r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', p_text) and 'Expedicion' not in p_text:
            if len(p.runs) >= 2:
                p.runs[0].text = fecha_hora_str
                p.runs[1].text = f". 849 0042 029 {tr_formatted}"
            else:
                p.text = f"{fecha_hora_str}. 849 0042 029 {tr_formatted}"
        elif p_text.startswith('RANG. AUT'):
            p.text = f"RANG. AUT {serie} 1 AL {serie} 1000000"
        elif 'Expedicion' in p_text and ':' in p_text:
            if len(p.runs) >= 3:
                p.runs[2].text = f": {fecha_hora_str}. 849"
            else:
                p.text = f"F. Expedicion: {fecha_hora_str}. 849"
        elif p_text.startswith('Cufe:'):
            p.text = f"Cufe:{cufe[:34]}"
            if idx + 1 < len(doc.paragraphs):
                doc.paragraphs[idx + 1].text = cufe[34:64]
        elif '>> MAPA:' in p_text:
            p.text = f">> MAPA:6536 ID: {fecha_compact}-{tr_formatted}768"
        elif p_text.startswith('NUTRX:'):
            p.text = f"NUTRX:{fecha_compact}05800008{tr_formatted}"

    # Inyección dinámica de productos en plantilla ÉXITO
    productos = factura.productos.all().order_by('id')
    if productos.exists():
        idx_plu = None
        idx_total_item = None
        for p_idx, p in enumerate(doc.paragraphs):
            if 'PLU' in p.text and 'DETALLE' in p.text:
                idx_plu = p_idx
            elif 'Total Item' in p.text:
                idx_total_item = p_idx
                break

        if idx_plu is not None and idx_total_item is not None:
            target_p = doc.paragraphs[idx_total_item]
            parrafos_a_borrar = [doc.paragraphs[k] for k in range(idx_plu + 1, idx_total_item)]
            for p in parrafos_a_borrar:
                p._p.getparent().remove(p._p)

            for idx, prod in enumerate(productos, 1):
                cod = prod.codigo or f"{random.randint(3100000, 3999999)}"
                lineas_prod = formatear_lineas_producto_exito(
                    idx=idx,
                    codigo=cod,
                    nombre=prod.nombre,
                    cantidad=prod.cantidad,
                    valor_unitario=prod.valor_unitario,
                    total_linea=prod.total_linea
                )
                for l in lineas_prod:
                    new_p = target_p.insert_paragraph_before(l)
                    new_p.style = doc.styles['Normal']

            target_p.text = f"Total Item : {productos.count()}"

        total_fmt = fmt_pesos(factura.total, separador_miles=".")
        subtotal_fmt = fmt_pesos(factura.subtotal or factura.total, separador_miles=".")
        iva_fmt = fmt_pesos(factura.iva, separador_miles=".")
        for p in doc.paragraphs:
            if p.text.startswith('SUBTOTAL'):
                p.text = f"SUBTOTAL                      {subtotal_fmt}"
            elif p.text.startswith('VALOR TOTAL'):
                p.text = f"VALOR TOTAL                   {total_fmt}"
            elif p.text.startswith('PAGA'):
                p.text = f"PAGA                          {total_fmt}"
            elif p.text.startswith('A-IVA='):
                p.text = f"A-IVA=19%{total_fmt} {subtotal_fmt} {iva_fmt}"

    limpiar_resaltados(doc)
    return doc

def rellenar_factura_guia_estatica(doc_path, factura):
    """
    Relleno dinamico e inteligente para facturas subidas manualmente como guia / modelo.
    Permite cargar 1 sola factura Word y reproducirla en lote con:
    - Folio y Serie consecutivos actualizados por factura
    - CUFE unico y dinamico
    - Codigo QR dinamico apuntando a la verificacion oficial en el dominio
    - Fecha y Hora de emision actualizadas segun la programacion del lote
    - Codigo de barras actualizado (si la plantilla usa barcode numerico y no es Jumbo)
    - Limpieza automatica de resaltados amarillos
    """
    doc = docx.Document(doc_path)
    base_url = os.environ.get('APP_DOMAIN') or os.environ.get('RENDER_EXTERNAL_URL') or 'https://catalogo-vpfe1.onrender.com'
    base_url = base_url.rstrip('/')

    cufe = factura.cufe or str(uuid.uuid4()).replace('-', '') + str(uuid.uuid4()).replace('-', '')
    serie = factura.serie or ""
    folio = factura.folio or ""
    numero_factura = f"{serie}{folio}".strip()

    fecha = factura.fecha_emision or timezone.localdate()
    hora = factura.hora_emision or timezone.localtime(timezone.now()).time()

    fecha_iso = fecha.strftime('%Y-%m-%d')
    hora_iso = hora.strftime('%H:%M:%S')
    fecha_hora_str = f"{fecha_iso} {hora_iso}"
    fecha_latam = fecha.strftime('%d/%m/%Y')
    fecha_corta = fecha.strftime('%d/ %m/ %y')
    hora_corta = hora.strftime('%H : %M')
    fecha_compact = fecha.strftime('%Y%m%d')

    # Detectar base_folio y base_serie de la factura_base si existe
    factura_base = getattr(factura.lote_generacion, 'factura_base', None) if factura.lote_generacion else None
    base_folio = str(factura_base.folio).strip() if factura_base and factura_base.folio else None
    base_serie = str(factura_base.serie).strip() if factura_base and factura_base.serie else None

    # 1. Reemplazo de Codigo QR
    qr_stream = generar_qr_stream(f"{base_url}/?cufe={cufe}")
    qr_bytes = qr_stream.getvalue()

    qr_rid_identified = None
    for i, p in enumerate(doc.paragraphs):
        if 'CUFE' in p.text.upper() or 'QR' in p.text.upper():
            for cand_p in doc.paragraphs[i:min(i + 4, len(doc.paragraphs))]:
                if 'drawing' in cand_p._p.xml or 'graphic' in cand_p._p.xml:
                    eids = re.findall(r'embed=[\x22\x27]([^\x22\x27]+)', cand_p._p.xml)
                    if eids:
                        qr_rid_identified = eids[0]
                        break
            if qr_rid_identified:
                break

    if qr_rid_identified and qr_rid_identified in doc.part.rels:
        doc.part.rels[qr_rid_identified].target_part._blob = qr_bytes
    else:
        for rId, rel in doc.part.rels.items():
            if 'image' in rel.target_ref or 'image' in getattr(rel.target_part, 'content_type', ''):
                try:
                    img = Image.open(io.BytesIO(rel.target_part._blob))
                    ratio = img.width / float(img.height or 1)
                    if 0.8 <= ratio <= 1.25 and img.width >= 100:
                        rel.target_part._blob = qr_bytes
                except Exception:
                    pass

    # 2. Reemplazo de Barcode (si existe y tiene ratio >= 2.0 y la plantilla no es JUMBO)
    is_jumbo = 'jumbo' in os.path.basename(doc_path).lower() or getattr(factura, 'plantilla', '') == 'JUMBO'
    if not is_jumbo and numero_factura:
        for rId, rel in doc.part.rels.items():
            if 'image' in rel.target_ref or 'image' in getattr(rel.target_part, 'content_type', ''):
                try:
                    img = Image.open(io.BytesIO(rel.target_part._blob))
                    ratio = img.width / float(img.height or 1)
                    if ratio >= 2.2:
                        bc_stream = generar_barcode_stream(numero_factura)
                        rel.target_part._blob = bc_stream.getvalue()
                        break
                except Exception:
                    pass

    def reemplazar_en_parrafo(parrafo, target_regex, replacement_text):
        if re.search(target_regex, parrafo.text):
            for run in parrafo.runs:
                if re.search(target_regex, run.text):
                    run.text = re.sub(target_regex, replacement_text, run.text)
                    return True
            if parrafo.runs:
                parrafo.runs[0].text = re.sub(target_regex, replacement_text, parrafo.text)
                for run in parrafo.runs[1:]:
                    run.text = ""
            else:
                parrafo.text = re.sub(target_regex, replacement_text, parrafo.text)
            return True
        return False

    # 3. Reemplazo en parrafos
    total_paragraphs = len(doc.paragraphs)
    for idx, p in enumerate(doc.paragraphs):
        p_text = p.text

        # CUFE
        if 'CUFE :' in p_text or 'CUFE:' in p_text or 'Cufe:' in p_text:
            m_hex = re.search(r'(?:CUFE\s*:\s*|Cufe:)([0-9a-fA-F]{20,})', p_text)
            if m_hex and idx + 1 < total_paragraphs and re.match(r'^[0-9a-fA-F]{15,}$', doc.paragraphs[idx + 1].text.strip()):
                p.text = f"Cufe:{cufe[:34]}"
                doc.paragraphs[idx + 1].text = cufe[34:64]
                continue
            elif m_hex:
                p.text = re.sub(r'(?:CUFE\s*:\s*|Cufe:)[0-9a-fA-F]{20,}', f"Cufe:{cufe}", p_text)
                continue
            elif idx + 1 < total_paragraphs:
                next_p = doc.paragraphs[idx + 1]
                if re.match(r'^[0-9a-fA-F]{30,}$', next_p.text.strip()) or len(next_p.text.strip()) > 25:
                    next_p.text = cufe
                    continue
        elif re.match(r'^[0-9a-fA-F]{40,96}$', p_text.strip()):
            p.text = cufe
            continue

        # Reemplazo de Folio y Serie
        if base_folio and base_folio in p_text:
            if base_serie and f"{base_serie}{base_folio}" in p_text:
                reemplazar_en_parrafo(p, re.escape(f"{base_serie}{base_folio}"), f"{serie}{folio}")
            elif base_serie and f"{base_serie}-{base_folio}" in p_text:
                reemplazar_en_parrafo(p, re.escape(f"{base_serie}-{base_folio}"), f"{serie}-{folio}")
            elif base_serie and f"{base_serie}- {base_folio}" in p_text:
                reemplazar_en_parrafo(p, re.escape(f"{base_serie}- {base_folio}"), f"{serie}- {folio}")
            elif base_serie and f"{base_serie}  {base_folio}" in p_text:
                reemplazar_en_parrafo(p, re.escape(f"{base_serie}  {base_folio}"), f"{serie}  {folio}")
            elif base_serie and f"{base_serie} {base_folio}" in p_text:
                reemplazar_en_parrafo(p, re.escape(f"{base_serie} {base_folio}"), f"{serie} {folio}")
            else:
                reemplazar_en_parrafo(p, re.escape(base_folio), folio)
        elif 'NUMERO:' in p_text.upper():
            reemplazar_en_parrafo(p, r'(NUMERO:\s*)([A-Za-z0-9\-_]+)', fr"\g<1>{numero_factura}")
        elif 'FACTURADOR  :' in p_text:
            reemplazar_en_parrafo(p, r'(FACTURADOR\s*:\s*)([A-Za-z0-9]+-\s*[0-9]+)', fr"\g<1>{serie}- {folio}")
        elif 'FAC.    ELECTRONICA   DE  VENTA' in p_text:
            reemplazar_en_parrafo(p, r'(FAC\.\s+ELECTRONICA\s+DE\s+VENTA\s+)([A-Za-z0-9]+\s+[0-9]+)', fr"\g<1>{serie}  {folio}")

        # Reemplazo de Fechas y Horas
        if 'Expedicion' in p_text and ':' in p_text:
            reemplazar_en_parrafo(p, r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', fecha_hora_str)
        elif re.search(r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', p_text):
            reemplazar_en_parrafo(p, r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', fecha_hora_str)
        elif re.search(r'\d{2}/\s*\d{2}/\s*\d{2,4}', p_text) and re.search(r'\d{2}\s*:\s*\d{2}', p_text):
            reemplazar_en_parrafo(p, r'\d{2}\s*:\s*\d{2}', hora_corta)
            reemplazar_en_parrafo(p, r'\d{2}/\s*\d{2}/\s*\d{2}', fecha_corta)
        elif '>> MAPA:' in p_text:
            reemplazar_en_parrafo(p, r'ID:\s*\d{8}-', f"ID: {fecha_compact}-")
        elif 'NUTRX:' in p_text:
            reemplazar_en_parrafo(p, r'NUTRX:\d{8}', f"NUTRX:{fecha_compact}")

    # 4. Procesar tablas si las hay
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                c_text = cell.text
                if base_folio and base_folio in c_text:
                    if base_serie and f"{base_serie}{base_folio}" in c_text:
                        cell.text = cell.text.replace(f"{base_serie}{base_folio}", f"{serie}{folio}")
                    elif base_serie and f"{base_serie}-{base_folio}" in c_text:
                        cell.text = cell.text.replace(f"{base_serie}-{base_folio}", f"{serie}-{folio}")
                    else:
                        cell.text = cell.text.replace(base_folio, folio)
                elif 'NUMERO:' in c_text.upper():
                    cell.text = re.sub(r'(NUMERO:\s*)([A-Za-z0-9\-_]+)', fr"\g<1>{numero_factura}", cell.text, flags=re.IGNORECASE)
                if 'CUFE' in c_text.upper():
                    m_hex = re.search(r'[0-9a-fA-F]{30,}', c_text)
                    if m_hex:
                        cell.text = cell.text.replace(m_hex.group(0), cufe)

    limpiar_resaltados(doc)
    return doc

class FormattedNumber(float):
    """
    Permite que los valores numericos se comporten como float/int para comparaciones en Jinja
    ({% if iva > 0 %}, {% if inc_total > 0 %}) y se impriman con formato de moneda/miles al interpolar ({{ iva }}).
    """
    def __new__(cls, val, fmt_str=None):
        try:
            f = float(val or 0)
        except Exception:
            f = 0.0
        obj = super().__new__(cls, f)
        obj.fmt_str = fmt_str
        return obj

    def __str__(self):
        if self.fmt_str:
            return self.fmt_str
        if self.is_integer():
            return f"{int(self):,}".replace(',', '.')
        return f"{self:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    def __format__(self, format_spec):
        return str(self)


class AutoSafeUndefined(jinja2.Undefined):
    """
    Evita cualquier excepcion por variables no definidas en plantillas Jinja/DocxTemplate.
    Soporta comparaciones seguras (> 0, < 0, ==), iteraciones vacias y strings vacios.
    """
    def _fail_with_undefined_error(self, *args, **kwargs):
        return ''
    def __str__(self):
        return ''
    def __html__(self):
        return ''
    def __bool__(self):
        return False
    def __iter__(self):
        return iter([])
    def __int__(self):
        return 0
    def __float__(self):
        return 0.0
    def __eq__(self, other):
        return False
    def __ne__(self, other):
        return True
    def __gt__(self, other):
        return False
    def __lt__(self, other):
        return False
    def __ge__(self, other):
        return False
    def __le__(self, other):
        return False
    def __getitem__(self, key):
        return self
    def __getattr__(self, name):
        return self
    def __call__(self, *args, **kwargs):
        return self


def construir_contexto_factura(factura, docxtpl_obj=None):
    base_url = os.environ.get('APP_DOMAIN') or os.environ.get('RENDER_EXTERNAL_URL') or 'https://catalogo-vpfe1.onrender.com'
    base_url = base_url.rstrip('/')
    qr_url = f"{base_url}/?cufe={factura.cufe}"

    qr_image = None
    bc_image = None
    if docxtpl_obj:
        try:
            qr_stream = generar_qr_stream(qr_url)
            qr_image = InlineImage(docxtpl_obj, qr_stream, width=Mm(35))
        except Exception:
            qr_image = ""
        try:
            num_factura = f"{factura.serie or ''}{factura.folio or ''}".strip()
            bc_stream = generar_barcode_stream(num_factura)
            bc_image = InlineImage(docxtpl_obj, bc_stream, width=Mm(45))
        except Exception:
            bc_image = ""

    # Preparar productos
    productos_list = []
    tot_cant = 0
    prods = factura.productos.all() if factura.pk else []
    for p in prods:
        c = p.cantidad or 1
        tot_cant += c
        pu = p.valor_unitario or Decimal('0')
        puv = getattr(p, 'precio_unitario_venta', None) or pu
        tl = p.total_linea or (c * pu)
        productos_list.append({
            'codigo': p.codigo or '',
            'nombre': p.nombre or '',
            'descripcion': getattr(p, 'descripcion', '') or p.nombre or '',
            'observacion': getattr(p, 'observacion', '') or '',
            'cantidad': FormattedNumber(c, str(int(c) if float(c).is_integer() else c)),
            'valor_unitario': FormattedNumber(pu, fmt_pesos(pu, separador_miles=".")),
            'precio_unitario': FormattedNumber(pu, fmt_pesos(pu, separador_miles=".")),
            'precio_unitario_venta': FormattedNumber(puv, fmt_pesos(puv, separador_miles=".")),
            'descuento': FormattedNumber(p.descuento or 0, fmt_pesos(p.descuento or 0, separador_miles=".")),
            'total_linea': FormattedNumber(tl, fmt_pesos(tl, separador_miles=".")),
            'total': FormattedNumber(tl, fmt_pesos(tl, separador_miles=".")),
            'tipo_iva': '19%',
        })

    subtotal_val = factura.subtotal or factura.total or Decimal('0')
    iva_val = factura.iva or Decimal('0')
    inc_val = getattr(factura, 'inc_total', None) or Decimal('0')
    total_val = factura.total or Decimal('0')
    desc_val = getattr(factura, 'descuento_detalle_total', None) or Decimal('0')
    tot_imp = getattr(factura, 'total_impuesto', None) or (iva_val + inc_val)

    fecha_emision_str = factura.fecha_emision.strftime('%d/%m/%Y') if factura.fecha_emision else ''
    fecha_emision_iso = factura.fecha_emision.strftime('%Y-%m-%d') if factura.fecha_emision else ''
    hora_emision_str = factura.hora_emision.strftime('%H:%M:%S') if factura.hora_emision else ''
    fecha_venc_str = factura.fecha_vencimiento.strftime('%d/%m/%Y') if factura.fecha_vencimiento else ''
    fecha_compact = factura.fecha_emision.strftime('%Y%m%d') if factura.fecha_emision else ''

    tr = getattr(factura, 'transaccion', '0197') or '0197'
    term = getattr(factura, 'terminal', '1034') or '1034'
    reg = getattr(factura, 'registro', '05') or '05'
    cajero = getattr(factura, 'cajero', 'CAJERO 01') or 'CAJERO 01'

    iva_items = [{
        'tipo': 'IVA',
        'tarifa': '19%',
        'compra': fmt_pesos(total_val, separador_miles="."),
        'base': fmt_pesos(subtotal_val, separador_miles="."),
        'impuesto': fmt_pesos(iva_val, separador_miles=".")
    }]

    context = {
        # Identificacion
        'serie': factura.serie or '',
        'folio': factura.folio or '',
        'numero': f"{factura.serie or ''}{factura.folio or ''}".strip(),
        'numero_factura': f"{factura.serie or ''}{factura.folio or ''}".strip(),
        'cufe': factura.cufe or '',
        'cufe_linea_1': factura.cufe[:32] if factura.cufe else '',
        'cufe_linea_2': factura.cufe[32:64] if factura.cufe and len(factura.cufe) >= 64 else '',
        'cufe_linea_3': factura.cufe[64:] if factura.cufe and len(factura.cufe) > 64 else '',

        # Fechas y Horas
        'fecha_emision': fecha_emision_str,
        'fecha_emision_iso': fecha_emision_iso,
        'fecha_vencimiento': fecha_venc_str,
        'hora_emision': hora_emision_str,
        'fecha_hora': f"{fecha_emision_iso} {hora_emision_str}",
        'fecha_expedicion_corta': fecha_emision_str,
        'fecha_hora_operacion': f"{fecha_emision_iso} {hora_emision_str}",
        'fecha_proceso': fecha_emision_str,
        'hora_proceso': hora_emision_str,
        'fecha_puntos': fecha_emision_str,

        # Emisor
        'emisor_nombre': factura.emisor_nombre or '',
        'emisor_nit': factura.emisor_nit or '',
        'emisor_direccion': factura.emisor_direccion or '',
        'emisor_telefono': factura.emisor_telefono or '',
        'emisor_correo': factura.emisor_correo or '',
        'almacen': getattr(factura, 'terminal', 'PRINCIPAL') or 'PRINCIPAL',
        'nombre_tienda': factura.emisor_nombre or '',
        'sede_linea': factura.emisor_direccion or '',

        # Receptor
        'receptor_nombre': factura.receptor_nombre or '',
        'receptor_nit': factura.receptor_nit or '',
        'receptor_nit_padded': str(factura.receptor_nit or '').rjust(13, '0'),
        'receptor_direccion': factura.receptor_direccion or '',
        'receptor_telefono': factura.receptor_telefono or '',
        'receptor_correo': factura.receptor_correo or '',
        'receptor_email': factura.receptor_correo or '',

        # Operacion y Caja
        'forma_pago': factura.forma_pago or 'CONTADO',
        'forma_pago_mayus': (factura.forma_pago or 'CONTADO').upper(),
        'medio_pago': factura.medio_pago or 'EFECTIVO',
        'cajero': cajero,
        'cajero_id': cajero[:10] if cajero else '01',
        'terminal': term,
        'registro': reg,
        'transaccion': tr,
        'maquina': term,
        'caja': term,
        'caja_registro': reg,
        'pedido': getattr(factura, 'orden_pedido', '1') or '1',
        'tipo_venta': 'VENTA EN LINEA',
        'tirilla_id': f"{fecha_compact}-{tr}",
        'software_pos': 'VPFE POS 3.0',
        'ciiu': '4711',
        'proveedor_factura': 'CADENA S.A.',

        # Totales numericos formateados y comparables
        'subtotal': FormattedNumber(subtotal_val, fmt_pesos(subtotal_val, separador_miles=".")),
        'subtotal_total_fmt': fmt_pesos(subtotal_val, separador_miles="."),
        'base_total_fmt': fmt_pesos(subtotal_val, separador_miles="."),
        'total': FormattedNumber(total_val, fmt_pesos(total_val, separador_miles=".")),
        'total_fmt': fmt_pesos(total_val, separador_miles="."),
        'compra_total_fmt': fmt_pesos(total_val, separador_miles="."),
        'iva': FormattedNumber(iva_val, fmt_pesos(iva_val, separador_miles=".")),
        'inc_total': FormattedNumber(inc_val, fmt_pesos(inc_val, separador_miles=".")),
        'inc_porcentaje': 8,
        'total_impuesto': FormattedNumber(tot_imp, fmt_pesos(tot_imp, separador_miles=".")),
        'impuesto_total_fmt': fmt_pesos(tot_imp, separador_miles="."),
        'impuesto_linea': fmt_pesos(tot_imp, separador_miles="."),
        'descuento_total': FormattedNumber(desc_val, fmt_pesos(desc_val, separador_miles=".")),
        'total_ahorros': '0',
        'ahorro_fmt': '0',
        'cambio': FormattedNumber(0, '0'),
        'cambio_fmt': '0',
        'pago_recibido': fmt_pesos(total_val, separador_miles="."),
        'pago_fmt': fmt_pesos(total_val, separador_miles="."),
        'puntos_cmr': '0',
        'puntos_total': '0',
        'total_articulos': len(productos_list),
        'total_lineas': len(productos_list),
        'total_lineas_un': tot_cant,

        # DIAN y Legales
        'numero_autorizacion': factura.numero_autorizacion or '',
        'rango_desde': factura.rango_desde or '',
        'rango_hasta': factura.rango_hasta or '',
        'vigencia': factura.vigencia or '',
        'vigencia_larga': factura.vigencia or '',
        'fecha_resolucion_larga': 'RESOLUCION DIAN VIGENTE',
        'resolucion_dian': factura.numero_autorizacion or '18764000001',
        'resolucion_gc': 'GRAN CONTRIBUYENTE',
        'xml_generado_por': factura.xml_generado_por or 'DIAN',
        'notas_legales': getattr(factura, 'texto_legal', '') or '',
        'notas_finales': getattr(factura, 'notas_finales', '') or '',
        'firma_linea': '========================================',
        'url_privacidad': 'https://www.dian.gov.co',

        # Listas de Productos e Impuestos
        'productos': productos_list,
        'iva_item': iva_items,

        # Elementos Graficos
        'qr': qr_image or '',
        'barcode': bc_image or '',
    }
    return context


def render_jinja_documento(tpl_path, factura):
    """
    Renderiza de forma 100% segura cualquier plantilla con sintaxis Jinja (docxtpl)
    sin fallos por variables faltantes o comparaciones numericas.
    """
    try:
        docxtpl_obj = DocxTemplate(tpl_path)
        context = construir_contexto_factura(factura, docxtpl_obj)
        jinja_env = jinja2.Environment(undefined=AutoSafeUndefined)
        docxtpl_obj.render(context, jinja_env=jinja_env)
        out_stream = io.BytesIO()
        docxtpl_obj.save(out_stream)
        out_stream.seek(0)
        doc = docx.Document(out_stream)
        limpiar_resaltados(doc)
        return doc
    except Exception as e:
        print(f"Aviso en render_jinja_documento: {e}. Usando relleno estatico de contingencia.")
        try:
            return rellenar_factura_guia_estatica(tpl_path, factura)
        except Exception as e2:
            print(f"Error en contingencia estatica: {e2}")
            doc = docx.Document(tpl_path)
            limpiar_resaltados(doc)
            return doc


def plantilla_tiene_jinja(doc_path):
    try:
        doc = docx.Document(doc_path)
        for p in doc.paragraphs:
            if "{{" in p.text and "}}" in p.text:
                return True
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    if "{{" in cell.text and "}}" in cell.text:
                        return True
        return False
    except Exception:
        return False

def generar_documento_factura(factura, template_path=None):
    """
    Genera y retorna un docx.Document listo para una factura individual.
    Selecciona automaticamente el motor de relleno segun la plantilla elegida.
    """
    tpl_path = template_path or resolver_plantilla_factura(factura)
    if not tpl_path or not os.path.exists(tpl_path):
        raise FileNotFoundError(f"No se encontro la plantilla para la factura {factura.serie}{factura.folio}")

    filename_lower = os.path.basename(tpl_path).lower()
    plantilla_tipo = getattr(factura, 'plantilla', '').upper()
    emisor_raw = (getattr(factura, 'emisor_nombre', '') or '').upper()
    nit_raw = (getattr(factura, 'emisor_nit', '') or '').replace('.', '').replace('-', '').strip()
    if plantilla_tipo == 'JUMBO' and ('EXITO' in emisor_raw or 'ÉXITO' in emisor_raw or '890900608' in nit_raw):
        plantilla_tipo = 'EXITO'
    elif plantilla_tipo == 'EXITO' and ('CENCOSUD' in emisor_raw or 'JUMBO' in emisor_raw or '900155107' in nit_raw):
        plantilla_tipo = 'JUMBO'

    # 1. Si el lote tiene plantilla personalizada subida por el usuario (factura_guia):
    if factura.lote_generacion and factura.lote_generacion.factura_guia and os.path.exists(factura.lote_generacion.factura_guia.path):
        if plantilla_tiene_jinja(tpl_path):
            return render_jinja_documento(tpl_path, factura)
        else:
            return rellenar_factura_guia_estatica(tpl_path, factura)

    # 2. Plantilla JUMBO
    if plantilla_tipo == 'JUMBO' or 'jumbo' in filename_lower:
        return rellenar_plantilla_jumbo(tpl_path, factura)

    # 3. Plantilla EXITO
    elif plantilla_tipo == 'EXITO' or 'exito' in filename_lower:
        return rellenar_plantilla_exito(tpl_path, factura)

    # 4. Cualquier otra plantilla con etiquetas Jinja (FRISBY, HOMECENTER, OLIMPICA, etc.)
    elif plantilla_tiene_jinja(tpl_path):
        return render_jinja_documento(tpl_path, factura)

    # 5. Plantilla estatica personalizada / OTRA
    else:
        return rellenar_factura_guia_estatica(tpl_path, factura)


def generar_factura_word_bytes(factura, template_path=None):
    """Retorna un BytesIO con el archivo .docx de la factura"""
    doc = generar_documento_factura(factura, template_path)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

def generar_lote_word_unificado_bytes(lote):
    """
    Combina todas las facturas de un lote en un unico archivo Word (.docx)
    separadas de forma continua en el formato de tirilla.
    Altamente tolerante a fallos: garantiza siempre la devolucion de un .docx valido.
    """
    facturas = list(lote.facturas.all().order_by('id'))
    
    # Si el lote no contiene facturas guardadas, intentar generar con su factura base
    if not facturas and lote.factura_base:
        try:
            from .admin import ejecutar_generacion_masiva_para_lote
            ejecutar_generacion_masiva_para_lote(lote)
            facturas = list(lote.facturas.all().order_by('id'))
        except Exception as e:
            print(f"Aviso al auto-generar facturas para lote #{lote.id}: {e}")

    # Si aun no hay facturas, usar la factura base directamente
    if not facturas:
        if lote.factura_base:
            doc = generar_documento_factura(lote.factura_base)
            stream = io.BytesIO()
            doc.save(stream)
            stream.seek(0)
            return stream
        else:
            raise ValueError(f"El lote #{lote.id} no contiene facturas generadas ni factura base asociada.")

    # Generar el documento para la primera factura exitosa
    doc_base = None
    factura_primera = None
    for f in facturas:
        try:
            doc_base = generar_documento_factura(f)
            factura_primera = f
            break
        except Exception as e:
            print(f"Error generando factura #{f.id}: {e}")
            continue

    if doc_base is None:
        raise ValueError("No se pudo procesar ninguna factura del lote.")

    # Combinar las siguientes facturas de forma segura
    otras_facturas = [f for f in facturas if f.id != factura_primera.id]
    if otras_facturas:
        try:
            from docxcompose.composer import Composer
            composer = Composer(doc_base)
            for f in otras_facturas:
                try:
                    doc_next = generar_documento_factura(f)
                    composer.append(doc_next)
                except Exception as ex:
                    print(f"Aviso: omitiendo factura #{f.id} en la composicion unificada: {ex}")
        except Exception as e:
            print(f"Aviso en Composer: {e}")

    # Guardar en stream con contingencia
    stream = io.BytesIO()
    try:
        doc_base.save(stream)
    except Exception as e:
        print(f"Error al guardar documento unificado ({e}), regenerando documento limpio...")
        doc_fresh = generar_documento_factura(factura_primera)
        stream = io.BytesIO()
        doc_fresh.save(stream)

    stream.seek(0)
    return stream


def generar_lote_zip_bytes(lote):
    """
    Genera un archivo comprimido .ZIP con:
    1. Todas las facturas individuales en .docx
    2. El archivo Word unificado con todas las facturas juntas.
    """
    facturas = lote.facturas.all().order_by('id')
    if not facturas.exists():
        raise ValueError("El lote no contiene facturas.")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Facturas individuales
        for f in facturas:
            filename = f"Factura_{f.serie}_{f.folio}.docx"
            f_stream = generar_factura_word_bytes(f)
            zip_file.writestr(filename, f_stream.getvalue())

        # 2. Documento unificado
        try:
            unificado_stream = generar_lote_word_unificado_bytes(lote)
            unificado_filename = f"Lote_{lote.id}_Facturas_Unificadas_{lote.factura_base.serie if lote.factura_base else 'POS'}.docx"
            zip_file.writestr(unificado_filename, unificado_stream.getvalue())
        except Exception as e:
            print("Error agregando unificado al zip:", e)

    zip_buffer.seek(0)
    return zip_buffer
