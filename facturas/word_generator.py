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
    Determina qué plantilla usar:
    1. Si el lote tiene factura_guia adjunta, usa esa.
    2. Si la factura tiene explícitamente el campo 'plantilla' ('JUMBO' o 'EXITO'), usa esa.
    3. Si coincide con alguno de los emisores registrados en plantillas_base, usa esa.
    4. Fallback a JUMBO.docx.
    """
    if factura.lote_generacion and factura.lote_generacion.factura_guia:
        if os.path.exists(factura.lote_generacion.factura_guia.path):
            return factura.lote_generacion.factura_guia.path

    plantillas_dir = os.path.join(settings.BASE_DIR, 'plantillas_base')

    # Chequear campo plantilla explícito
    plantilla_val = getattr(factura, 'plantilla', None)
    if plantilla_val in ['JUMBO', 'EXITO']:
        p = os.path.join(plantillas_dir, f"{plantilla_val}.docx")
        if os.path.exists(p):
            return p

    emisor_norm = normalizar_texto(factura.emisor_nombre)

    # Búsqueda por alias
    for nombre_plantilla, alias_list in PLANTILLA_ALIASES.items():
        for alias in alias_list:
            if normalizar_texto(alias) in emisor_norm:
                p = os.path.join(plantillas_dir, f"{nombre_plantilla}.docx")
                if os.path.exists(p):
                    return p

    # Fallback por defecto
    for fb in ['JUMBO.docx', 'EXITO.docx', 'defecto.docx']:
        fb_path = os.path.join(plantillas_dir, fb)
        if os.path.exists(fb_path):
            return fb_path

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
    """Elimina el fondo amarillo de todos los runs en el documento para impresión profesional."""
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

    bc_stream = generar_barcode_stream(rem)
    bc_bytes = bc_stream.getvalue()
    if 'rId9' in doc.part.rels:
        doc.part.rels['rId9'].target_part._blob = bc_bytes

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

    bc_stream = generar_barcode_stream(numero_factura)
    bc_bytes = bc_stream.getvalue()
    if 'rId7' in doc.part.rels:
        doc.part.rels['rId7'].target_part._blob = bc_bytes

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
    Relleno genérico de fallback para documentos guía subidos manualmente por el usuario.
    """
    doc = docx.Document(doc_path)
    base_url = os.environ.get('APP_DOMAIN') or os.environ.get('RENDER_EXTERNAL_URL') or 'https://catalogo-vpfe1.onrender.com'
    qr_url = f"{base_url.rstrip('/')}/?cufe={factura.cufe}"
    qr_png_bytes = generar_qr_stream(qr_url).getvalue()

    cufe_reemplazado = False
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        if 'CUFE :' in text or 'CUFE:' in text:
            if i + 1 < len(doc.paragraphs):
                next_p = doc.paragraphs[i + 1]
                if len(next_p.text.strip()) > 20:
                    next_p.text = factura.cufe
                    cufe_reemplazado = True

        if cufe_reemplazado and ('drawing' in p._p.xml or 'graphic' in p._p.xml):
            embed_ids = re.findall(r'embed=[\x22\x27]([^\x22\x27]+)', p._p.xml)
            for rId in embed_ids:
                if rId in doc.part.rels:
                    target_part = doc.part.rels[rId].target_part
                    if 'image' in target_part.content_type:
                        target_part._blob = qr_png_bytes
                        break

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
    Selecciona automáticamente el motor de relleno según la plantilla elegida.
    """
    tpl_path = template_path or resolver_plantilla_factura(factura)
    if not tpl_path or not os.path.exists(tpl_path):
        raise FileNotFoundError(f"No se encontró la plantilla para la factura {factura.serie}{factura.folio}")

    filename_lower = os.path.basename(tpl_path).lower()
    plantilla_tipo = getattr(factura, 'plantilla', '').upper()

    if plantilla_tipo == 'JUMBO' or 'jumbo' in filename_lower:
        return rellenar_plantilla_jumbo(tpl_path, factura)
    elif plantilla_tipo == 'EXITO' or 'exito' in filename_lower:
        return rellenar_plantilla_exito(tpl_path, factura)
    elif plantilla_tiene_jinja(tpl_path):
        docxtpl_obj = DocxTemplate(tpl_path)
        context = {
            'serie': factura.serie,
            'folio': factura.folio,
            'cufe': factura.cufe,
            'fecha_emision': factura.fecha_emision.strftime('%d/%m/%Y') if factura.fecha_emision else '',
            'hora_emision': factura.hora_emision.strftime('%H:%M:%S') if factura.hora_emision else '',
            'total': fmt_pesos(factura.total),
            'subtotal': fmt_pesos(factura.subtotal or factura.total),
            'iva': fmt_pesos(factura.iva),
        }
        docxtpl_obj.render(context)
        out_stream = io.BytesIO()
        docxtpl_obj.save(out_stream)
        out_stream.seek(0)
        doc = docx.Document(out_stream)
        limpiar_resaltados(doc)
        return doc
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
    Combina todas las facturas de un lote en un único archivo Word (.docx)
    separadas de forma continua exactamente en el formato de tirilla solicitado.
    """
    facturas = lote.facturas.all().order_by('id')
    if not facturas.exists():
        raise ValueError("El lote no contiene facturas.")

    doc_base = generar_documento_factura(facturas[0])
    if len(facturas) > 1:
        try:
            from docxcompose.composer import Composer
            composer = Composer(doc_base)
            for f in facturas[1:]:
                doc_next = generar_documento_factura(f)
                composer.append(doc_next)
        except Exception as e:
            print("Composer error:", e)

    stream = io.BytesIO()
    doc_base.save(stream)
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
