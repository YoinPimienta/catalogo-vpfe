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

    # Si se editaron productos en la factura base, actualizamos totales
    productos = factura.productos.all()
    if productos.exists() and (factura.total != Decimal('43650.00') or productos.count() > 1):
        total_fmt = fmt_pesos(factura.total)
        subtotal_fmt = fmt_pesos(factura.subtotal or factura.total)
        iva_fmt = fmt_pesos(factura.iva)
        cambio_fmt = fmt_pesos(Decimal('50000.00') - factura.total) if factura.total < Decimal('50000.00') else '0'

        if len(doc.paragraphs) > 18:
            doc.paragraphs[18].text = f"                   SUBTOTAL/TOTAL                                  {total_fmt}"
        if len(doc.paragraphs) > 21:
            doc.paragraphs[21].text = f"CAMBIO                                                                                 {cambio_fmt}"
        if len(doc.paragraphs) > 26:
            doc.paragraphs[26].text = f"N=19%                                  {subtotal_fmt}                       {iva_fmt}"
        if len(doc.paragraphs) > 29:
            doc.paragraphs[29].text = f" TOTAL                              {total_fmt}                           {iva_fmt}"

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

    # Si se editaron productos en la factura base, actualizamos totales
    productos = factura.productos.all()
    if productos.exists() and (factura.total != Decimal('10503202.00') or productos.count() != 3):
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
