import random
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
from docx.shared import Mm, Pt, Inches
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

    # Chequear campo plantilla explicito con validacion inteligente frente al emisor
    plantilla_val = getattr(factura, 'plantilla', None)
    em_norm = normalizar_texto(getattr(factura, 'emisor_nombre', ''))
    nit_norm = (getattr(factura, 'emisor_nit', '') or '').replace('.', '').replace('-', '').strip()
    serie_norm = (getattr(factura, 'serie', '') or '').upper()

    if 'EXITO' in em_norm or '890900608' in nit_norm or 'WJ53' in serie_norm:
        plantilla_val = 'EXITO'
    elif 'CENCOSUD' in em_norm or 'JUMBO' in em_norm or '900155107' in nit_norm or 'AU06' in serie_norm:
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



def generar_barcode_bytes_resized(code_text, target_width, target_height):
    """
    Genera un codigo de barras Code128 y lo redimensiona al tamano exacto del
    placeholder original en el documento Word, para evitar desbordamiento de pagina.
    target_width y target_height se expresan en pixeles.
    """
    stream = generar_barcode_stream(code_text)
    img = Image.open(stream)
    img_resized = img.resize((target_width, target_height), Image.LANCZOS)
    out = io.BytesIO()
    img_resized.save(out, format='PNG')
    return out.getvalue()


def _forzar_fuente_8pt(run):
    """
    Fuerza Calibri 8pt (sz=16 half-points) en el run dado, eliminando
    cualquier herencia de tamano mayor proveniente del estilo de la plantilla.
    """
    from docx.oxml.ns import qn
    from lxml import etree
    run.font.name = 'Calibri'
    run.font.size = Pt(8)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), 'Calibri')
    rFonts.set(qn('w:hAnsi'), 'Calibri')
    rFonts.set(qn('w:cs'), 'Calibri')
    for tag in ('w:sz', 'w:szCs'):
        elem = rPr.find(qn(tag))
        if elem is None:
            elem = etree.SubElement(rPr, qn(tag))
        elem.set(qn('w:val'), '16')

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
    qr_stream = generar_qr_stream(f"{base_url}/factura/{cufe}/")
    qr_bytes = qr_stream.getvalue()
    if 'rId8' in doc.part.rels:
        doc.part.rels['rId8'].target_part._blob = qr_bytes

    # Inyeccion dinamica de codigo de barras Code 128 nitido para REM
    # El placeholder original en JUMBO.docx mide 508x92 px; resizamos para no desbordar pagina
    try:
        bc_bytes = generar_barcode_bytes_resized(rem, target_width=508, target_height=92)
        if 'rId9' in doc.part.rels:
            doc.part.rels['rId9'].target_part._blob = bc_bytes
    except Exception as e:
        print(f"Aviso al inyectar codigo de barras Jumbo: {e}")

    # Reemplazo de CUFE y FACTURADOR con eliminacion limpia de nodos XML para no dejar saltos ni espacios extra
    parrafos_a_remover_jumbo = []
    for idx_p, p in enumerate(doc.paragraphs):
        p_text = p.text.strip()
        if p_text in ('CUFE :', 'CUFE : ', 'CUFE:'):
            if idx_p + 1 < len(doc.paragraphs):
                p_cufe = doc.paragraphs[idx_p + 1]
                p_cufe.text = cufe
                for r in p_cufe.runs:
                    _forzar_fuente_8pt(r)
            # Marcar parrafos residuales antes del QR para removerlos del XML
            for sub_i in range(idx_p + 2, min(idx_p + 5, len(doc.paragraphs))):
                p_sub = doc.paragraphs[sub_i]
                if 'drawing' in p_sub._p.xml or 'graphic' in p_sub._p.xml or 'FACTURADOR' in p_sub.text:
                    break
                parrafos_a_remover_jumbo.append(p_sub)
        elif 'FACTURADOR' in p_text and ('CUFE' in p_text or ':' in p_text):
            p.text = f"  FACTURADOR  :  {serie}- {folio}   - CUFE "
            for r in p.runs:
                _forzar_fuente_8pt(r)
            p_cufe_colocado = False
            for next_idx in range(idx_p + 1, min(idx_p + 7, len(doc.paragraphs))):
                p_next = doc.paragraphs[next_idx]
                txt_next = p_next.text.strip()
                if '----' in txt_next or 'Cadena' in txt_next or 'RESOL' in txt_next or 'FAC.' in txt_next:
                    break
                if not p_cufe_colocado:
                    p_next.text = cufe
                    for r in p_next.runs:
                        _forzar_fuente_8pt(r)
                    p_cufe_colocado = True
                else:
                    # Remover linea residual del CUFE viejo para no dejar lineas en blanco
                    parrafos_a_remover_jumbo.append(p_next)

    for p_del in parrafos_a_remover_jumbo:
        try:
            p_del._element.getparent().remove(p_del._element)
        except Exception:
            p_del.text = ""

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
    qr_stream = generar_qr_stream(f"{base_url}/factura/{cufe}/")
    qr_bytes = qr_stream.getvalue()
    if 'rId6' in doc.part.rels:
        doc.part.rels['rId6'].target_part._blob = qr_bytes

    # Inyeccion dinamica de codigo de barras Code 128 nitido para numero de factura
    # El placeholder original en EXITO.docx mide 988x200 px; resizamos para no desbordar pagina
    try:
        bc_bytes = generar_barcode_bytes_resized(numero_factura, target_width=988, target_height=200)
        if 'rId7' in doc.part.rels:
            doc.part.rels['rId7'].target_part._blob = bc_bytes
    except Exception as e:
        print(f"Aviso al inyectar codigo de barras Exito: {e}")

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

    # 1. Reemplazo de Codigo QR garantizado
    qr_stream = generar_qr_stream(f"{base_url}/factura/{cufe}/")
    qr_bytes = qr_stream.getvalue()

    qr_rid_identified = None
    # A. Buscar alrededor de palabras clave (CUFE o QR) en un rango amplio (-3 a +7)
    for i, p in enumerate(doc.paragraphs):
        p_txt_u = p.text.upper()
        if 'CUFE' in p_txt_u or 'QR' in p_txt_u or 'FACTURADOR' in p_txt_u:
            rango_cand = doc.paragraphs[max(0, i - 3):min(i + 8, len(doc.paragraphs))]
            for cand_p in rango_cand:
                if 'drawing' in cand_p._p.xml or 'blip' in cand_p._p.xml or 'imagedata' in cand_p._p.xml:
                    eids = re.findall(r'(?:embed|id)=[\x22\x27](rId\d+)', cand_p._p.xml)
                    for eid in eids:
                        if eid in doc.part.rels:
                            rel = doc.part.rels[eid]
                            if 'image' in rel.target_ref or 'image' in getattr(rel.target_part, 'content_type', ''):
                                try:
                                    img = Image.open(io.BytesIO(rel.target_part._blob))
                                    ratio = img.width / float(img.height or 1)
                                    if 0.75 <= ratio <= 1.3:
                                        qr_rid_identified = eid
                                        break
                                except Exception:
                                    pass
                    if qr_rid_identified:
                        break
        if qr_rid_identified:
            break

    # B. Si no se identifico cerca de CUFE, buscar cualquier imagen cuadrada en el documento
    if not qr_rid_identified:
        for rId, rel in doc.part.rels.items():
            if 'image' in rel.target_ref or 'image' in getattr(rel.target_part, 'content_type', ''):
                try:
                    img = Image.open(io.BytesIO(rel.target_part._blob))
                    ratio = img.width / float(img.height or 1)
                    if 0.75 <= ratio <= 1.3:
                        qr_rid_identified = rId
                        break
                except Exception:
                    pass

    # C. Inyectar bytes de QR en la relacion encontrada
    if qr_rid_identified and qr_rid_identified in doc.part.rels:
        doc.part.rels[qr_rid_identified].target_part._blob = qr_bytes
    else:
        # D. Si la plantilla no tiene imagen de QR, insertar imagen en placeholder o tras CUFE
        qr_insertado = False
        for p in doc.paragraphs:
            if '{{ qr }}' in p.text.lower() or '[qr]' in p.text.lower() or '[[qr]]' in p.text.lower():
                p.text = ""
                run = p.add_run()
                qr_stream.seek(0)
                run.add_picture(qr_stream, width=Inches(1.5), height=Inches(1.5))
                p.alignment = 1
                qr_insertado = True
                break
        if not qr_insertado:
            for i, p in enumerate(doc.paragraphs):
                if 'CUFE' in p.text.upper():
                    idx_ins = min(i + 2, len(doc.paragraphs) - 1)
                    target_p = doc.paragraphs[idx_ins]
                    new_p = target_p.insert_paragraph_before()
                    new_p.alignment = 1
                    qr_stream.seek(0)
                    new_p.add_run().add_picture(qr_stream, width=Inches(1.5), height=Inches(1.5))
                    break

    # 2. Reemplazo de Barcode (preservando dimensiones exactas sin deformar)
    is_jumbo = 'jumbo' in os.path.basename(doc_path).lower() or getattr(factura, 'plantilla', '') == 'JUMBO'
    if not is_jumbo and numero_factura:
        for rId, rel in doc.part.rels.items():
            if rId == qr_rid_identified:
                continue
            if 'image' in rel.target_ref or 'image' in getattr(rel.target_part, 'content_type', ''):
                try:
                    img = Image.open(io.BytesIO(rel.target_part._blob))
                    ratio = img.width / float(img.height or 1)
                    if ratio >= 2.0:
                        bc_bytes = generar_barcode_bytes_resized(numero_factura, target_width=img.width, target_height=img.height)
                        rel.target_part._blob = bc_bytes
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

    # 3. Reemplazo en parrafos (CUFE universal, Serie, Folio, Fechas y Horas)
    parrafos_a_remover = []
    total_paragraphs = len(doc.paragraphs)
    for idx, p in enumerate(doc.paragraphs):
        p_text = p.text
        p_clean = p_text.strip()
        p_upper = p_clean.upper()

        # Caso A: Estilo Éxito - "Cufe:..." o "CUFE:..." (con o sin hash pegado)
        if re.match(r'^(?:Cufe|CUFE)\s*:', p_clean):
            p.text = f"Cufe:{cufe[:34]}"
            for r in p.runs:
                _forzar_fuente_8pt(r)
            if idx + 1 < total_paragraphs:
                p_next = doc.paragraphs[idx + 1]
                txt_next = p_next.text.strip()
                if re.match(r'^[0-9a-zA-Z\s]{15,90}$', txt_next) and not any(k in txt_next for k in ('Toshiba', 'SuperMarket', 'Cadena', 'RESOL', 'NIT', 'DIAN', 'RANG', 'FACTURA')):
                    p_next.text = cufe[34:]
                    for r in p_next.runs:
                        _forzar_fuente_8pt(r)
            continue

        # Caso B: Estilo Jumbo / Frisby - "CUFE :" o "CUFE:" en su propia linea
        elif p_clean in ('CUFE :', 'CUFE : ', 'CUFE:', 'CUFE'):
            if idx + 1 < total_paragraphs:
                p_next = doc.paragraphs[idx + 1]
                p_next.text = cufe
                for r in p_next.runs:
                    _forzar_fuente_8pt(r)
            for sub_i in range(idx + 2, min(idx + 6, total_paragraphs)):
                p_sub = doc.paragraphs[sub_i]
                if 'drawing' in p_sub._p.xml or 'graphic' in p_sub._p.xml or 'FACTURADOR' in p_sub.text:
                    break
                parrafos_a_remover.append(p_sub)
            continue

        # Caso C: Estilo FACTURADOR Jumbo - "FACTURADOR : ... - CUFE"
        elif 'FACTURADOR' in p_upper and ('CUFE' in p_upper or ':' in p_text):
            p.text = f"  FACTURADOR  :  {serie}- {folio}   - CUFE "
            for r in p.runs:
                _forzar_fuente_8pt(r)
            p_cufe_colocado = False
            for next_idx in range(idx + 1, min(idx + 8, total_paragraphs)):
                p_next = doc.paragraphs[next_idx]
                txt_next = p_next.text.strip()
                if any(k in txt_next for k in ('----', 'Cadena', 'RESOL', 'FAC.', 'VIGENCIA')):
                    break
                if not p_cufe_colocado:
                    p_next.text = cufe
                    for r in p_next.runs:
                        _forzar_fuente_8pt(r)
                    p_cufe_colocado = True
                else:
                    parrafos_a_remover.append(p_next)
            continue

        # Caso D: Parrafo que contiene "CUFE" y un hash largo (30+ caracteres hex)
        elif 'CUFE' in p_upper and re.search(r'[0-9a-fA-F]{30,}', p_text):
            p.text = re.sub(r'[0-9a-fA-F]{30,}', cufe, p_text)
            for r in p.runs:
                _forzar_fuente_8pt(r)
            continue

        # Caso E: Placeholder Jinja en plantilla
        elif '{{' in p_text and ('cufe' in p_text.lower() or 'qr' in p_text.lower()):
            reemplazar_en_parrafo(p, r'\{\{\s*cufe\s*\}\}', cufe)
            reemplazar_en_parrafo(p, r'\{\{\s*cufe_linea_1\s*\}\}', cufe[:34])
            reemplazar_en_parrafo(p, r'\{\{\s*cufe_linea_2\s*\}\}', cufe[34:])
            continue

        # Caso F: Residuo de hash solitario tras FACTURADOR o CUFE
        elif re.match(r'^[0-9a-zA-Z\s]{25,120}$', p_clean) and any(k in doc.paragraphs[max(0, idx-2)].text for k in ('CUFE', 'FACTURADOR', 'Cufe')):
            parrafos_a_remover.append(p)
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

    for p_del in parrafos_a_remover:
        try:
            p_del._element.getparent().remove(p_del._element)
        except Exception:
            p_del.text = ""

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
    qr_url = f"{base_url}/factura/{factura.cufe}/"

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



def extraer_datos_de_factura_guia_docx(doc_path):
    """
    Lee una factura modelo (.docx) adjunta por el cliente y extrae de forma inteligente:
    - Lista de productos con codigo, nombre, cantidad, valor unitario y total
    - Totales: total, subtotal, iva
    Permite que facturas personalizadas se reflejen 100% fieles en la base de datos y DIAN.
    """
    if not doc_path or not os.path.exists(doc_path):
        return {'productos': [], 'total': None, 'subtotal': None, 'iva': None}

    try:
        doc = docx.Document(doc_path)
    except Exception as e:
        print(f"Aviso: no se pudo abrir {doc_path} para extraer datos: {e}")
        return {'productos': [], 'total': None, 'subtotal': None, 'iva': None}

    textos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    
    total = None
    subtotal = None
    iva = None
    
    for t in textos:
        t_up = t.upper()
        if t_up.startswith('VALOR TOTAL') or t_up.startswith('TOTAL :') or t_up.startswith('TOTAL:'):
            m = re.search(r'([\d\.,]{3,})', t[11:])
            if m and total is None:
                total = limpiar_monto_colombiano(m.group(1))
        elif 'SUBTOTAL/TOTAL' in t_up:
            m = re.search(r'([\d\.,]{3,})', t)
            if m and total is None:
                total = limpiar_monto_colombiano(m.group(1))
        elif t_up.startswith('TOTAL') and not t_up.startswith('TOTAL ITEM') and total is None:
            m = re.search(r'([\d\.,]{3,})', t[5:])
            if m:
                total = limpiar_monto_colombiano(m.group(1))
        
        if t_up.startswith('SUBTOTAL') and 'SUBTOTAL/TOTAL' not in t_up:
            m = re.search(r'([\d\.,]{3,})', t[8:])
            if m and subtotal is None:
                subtotal = limpiar_monto_colombiano(m.group(1))
        
        if 'A-IVA=' in t_up:
            partes = re.findall(r'[\d\.,]{3,}', t)
            if len(partes) >= 3:
                if total is None:
                    total = limpiar_monto_colombiano(partes[0])
                if subtotal is None:
                    subtotal = limpiar_monto_colombiano(partes[1])
                if iva is None:
                    iva = limpiar_monto_colombiano(partes[2])
        elif 'N=19%' in t_up or 'IVA=' in t_up:
            partes = re.findall(r'[\d\.,]{3,}', t)
            if len(partes) >= 2:
                if subtotal is None:
                    subtotal = limpiar_monto_colombiano(partes[0])
                if iva is None:
                    iva = limpiar_monto_colombiano(partes[1])

    productos = []
    # 1. Buscar formato Exito: PLU DETALLE PRECIO ... Total Item
    idx_plu = None
    idx_total_item = None
    for idx, p in enumerate(doc.paragraphs):
        txt = p.text.upper()
        if 'PLU' in txt and 'DETALLE' in txt:
            idx_plu = idx
        elif 'TOTAL ITEM' in txt and idx_plu is not None:
            idx_total_item = idx
            break
            
    if idx_plu is not None and idx_total_item is not None:
        p_lines = [doc.paragraphs[k].text.strip() for k in range(idx_plu + 1, idx_total_item) if doc.paragraphs[k].text.strip()]
        i = 0
        while i < len(p_lines):
            l = p_lines[i]
            m_header = re.match(r'^\d+\s+(\d+)/u\s+[xX]\s*([\$\d\.,]+)', l)
            if m_header:
                cant = int(m_header.group(1))
                precio = limpiar_monto_colombiano(m_header.group(2))
                i += 1
                desc_parts = []
                while i < len(p_lines) and not re.match(r'^\d+\s+\d+/u', p_lines[i]):
                    desc_parts.append(p_lines[i])
                    i += 1
                desc_full = ' '.join(desc_parts)
                m_cod = re.match(r'^(\d{5,10})\s+(.*)', desc_full)
                if m_cod:
                    cod = m_cod.group(1)
                    nom = re.sub(r'[\$\d\.,]+[A-Za-z]?\s*$', '', m_cod.group(2)).strip()
                else:
                    cod = "3100827"
                    nom = re.sub(r'[\$\d\.,]+[A-Za-z]?\s*$', '', desc_full).strip()
                productos.append({
                    'codigo': cod,
                    'nombre': nom or 'Producto',
                    'cantidad': cant,
                    'valor_unitario': precio,
                    'total_linea': precio * cant
                })
            else:
                i += 1

    # 2. Buscar formato Jumbo si no hubo Exito
    if not productos:
        for idx, p in enumerate(doc.paragraphs):
            txt = p.text.strip()
            m_jumbo = re.match(r'^(\d+)\s+[xX]\s*([\$\d\.,]+)', txt)
            if m_jumbo and idx + 1 < len(doc.paragraphs):
                cant = int(m_jumbo.group(1))
                precio = limpiar_monto_colombiano(m_jumbo.group(2))
                next_txt = doc.paragraphs[idx + 1].text.strip()
                m_next = re.match(r'^(\d+)\s+(.*?)(?:\s+\$([\d\.,]+))?$', next_txt)
                if m_next:
                    cod = m_next.group(1)
                    nom = m_next.group(2).strip()
                else:
                    cod = "3669368"
                    nom = next_txt
                productos.append({
                    'codigo': cod,
                    'nombre': nom,
                    'cantidad': cant,
                    'valor_unitario': precio,
                    'total_linea': precio * cant
                })

    return {
        'productos': productos,
        'total': total,
        'subtotal': subtotal,
        'iva': iva
    }

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
    em_norm = normalizar_texto(getattr(factura, 'emisor_nombre', ''))
    nit_norm = (getattr(factura, 'emisor_nit', '') or '').replace('.', '').replace('-', '').strip()
    serie_norm = (getattr(factura, 'serie', '') or '').upper()

    if 'EXITO' in em_norm or '890900608' in nit_norm or 'WJ53' in serie_norm:
        plantilla_tipo = 'EXITO'
    elif 'CENCOSUD' in em_norm or 'JUMBO' in em_norm or '900155107' in nit_norm or 'AU06' in serie_norm:
        plantilla_tipo = 'JUMBO'

    # 1. Si el lote tiene plantilla personalizada subida por el usuario (factura_guia):
    if factura.lote_generacion and factura.lote_generacion.factura_guia and os.path.exists(factura.lote_generacion.factura_guia.path):
        if plantilla_tiene_jinja(tpl_path):
            return render_jinja_documento(tpl_path, factura)
        # Respetar 100% los productos y diseno organizados por el cliente en su archivo modelo .docx
        # rellenar_factura_guia_estatica preserva los productos exactos y solo actualiza folio, serie, CUFE, QR, barcode, fecha y hora
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
    Optimizado con control de memoria y tiempo para Render.
    """
    import gc, time
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

    otras_facturas = [f for f in facturas if f.id != factura_primera.id]
    
    # Cap de seguridad ampliado a 100 facturas para el Word unificado (cubre pedidos de 50 facturas)
    if len(otras_facturas) > 100:
        otras_facturas = otras_facturas[:100]

    if otras_facturas:
        try:
            from docxcompose.composer import Composer
            composer = Composer(doc_base)
            t_start = time.time()
            for idx, f in enumerate(otras_facturas):
                # Limite preventivo ampliado a 140 segundos (aprovecha timeout de 180s de Gunicorn)
                if time.time() - t_start > 140.0:
                    print(f"Aviso: limite de tiempo alcanzado al componer lote #{lote.id} ({idx+1} facturas unificadas)")
                    break
                try:
                    doc_next = generar_documento_factura(f)
                    composer.append(doc_next)
                    del doc_next
                except Exception as ex:
                    print(f"Aviso: omitiendo factura #{f.id} en la composicion unificada: {ex}")
                
                if (idx + 1) % 5 == 0:
                    gc.collect()
        except Exception as e:
            print(f"Aviso en Composer: {e}")

    gc.collect()
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
    2. El archivo Word unificado con todas las facturas juntas (para lotes <= 30 facturas).
    """
    import gc
    facturas = list(lote.facturas.all().order_by('id'))
    if not facturas:
        raise ValueError("El lote no contiene facturas.")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Facturas individuales
        for idx, f in enumerate(facturas):
            filename = f"Factura_{f.serie}_{f.folio}.docx"
            f_stream = generar_factura_word_bytes(f)
            zip_file.writestr(filename, f_stream.getvalue())
            del f_stream
            if (idx + 1) % 10 == 0:
                gc.collect()

        # 2. Documento unificado protegido (soporta hasta 100 facturas en lote)
        if len(facturas) <= 100:
            try:
                unificado_stream = generar_lote_word_unificado_bytes(lote)
                unificado_filename = f"Lote_{lote.id}_Facturas_Unificadas_{lote.factura_base.serie if lote.factura_base else 'POS'}.docx"
                zip_file.writestr(unificado_filename, unificado_stream.getvalue())
                del unificado_stream
            except Exception as e:
                print(f"Aviso agregando unificado al zip: {e}")

    gc.collect()
    zip_buffer.seek(0)
    return zip_buffer
