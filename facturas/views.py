from django.shortcuts import render
from django.template.loader import render_to_string
from django.http import HttpResponse, JsonResponse
from django.conf import settings
from .models import Factura
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import os
import base64

def buscar_factura(request):
    cufe = request.GET.get('cufe') or request.POST.get('cufe')
    error = None
    factura = None

    if request.method == 'POST' and cufe:
        cufe_clean = str(cufe).strip()
        try:
            factura = Factura.objects.get(cufe=cufe_clean)
            if factura.estado == 'Anulada' or (factura.lote_generacion and factura.lote_generacion.estado == 'DESACTIVADO'):
                error = "Este documento ha sido anulado o desactivado y no está disponible para búsquedas normales."
                factura = None
            else:
                return render(request, 'factura_detalle.html', {'factura': factura})
        except Factura.DoesNotExist:
            error = "Documento no encontrado. Verifique el código ingresado e intente nuevamente."
        except Exception as e:
            print("Error interno al buscar factura:", e)
            error = "Ocurrió un error inesperado al procesar la búsqueda."

    return render(request, 'buscar.html', {'error': error, 'cufe': cufe})
def ver_factura(request, cufe):
    try:
        factura = Factura.objects.get(cufe=cufe)
        if factura.estado == 'Anulada' or (factura.lote_generacion and factura.lote_generacion.estado == 'DESACTIVADO'):
            return render(request, 'buscar.html', {'error': "Este documento ha sido anulado o desactivado y no está disponible."})
        return render(request, 'factura_detalle.html', {'factura': factura})
    except Factura.DoesNotExist:
        return render(request, 'buscar.html', {'error': "Documento no encontrado."})

def descargar_pdf_factura(request, cufe):
    try:
        from weasyprint import HTML
        from django.db.models import F
        factura = Factura.objects.get(cufe=cufe)
        
        # Registrar descarga para auditoría
        Factura.objects.filter(pk=factura.pk).update(descargas=F('descargas') + 1)
        
        factura.recalcular_totales()
        factura.generar_qr()
        html_string = render_to_string('factura_pdf.html', {'factura': factura, 'request': request})
        
        # Generar el PDF
        pdf_file = HTML(string=html_string, base_url=request.build_absolute_uri('/')).write_pdf()
        
        current_serie = factura.serie if factura.serie else "I0"
        current_folio = factura.folio if factura.folio else "00"
        
        response = HttpResponse(pdf_file, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="Factura_{current_serie}_{current_folio}.pdf"'
        return response
    except Factura.DoesNotExist:
        return HttpResponse("Factura no encontrada.", status=404)
    except Exception as e:
        import traceback
        error_msg = f"Error generando PDF: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_msg)
        return HttpResponse(error_msg, status=500, content_type="text/plain")

import unicodedata

PLANTILLA_ALIASES = {
    "EXITO": ["EXITO", "ALMACENES EXITO", "ALMACENES EXITO SA", "ALMACENES EXITO S.A"],
    "FRISBY": ["FRISBY", "FRISBY SA", "FRISBY S.A", "FRISBY S.A. BIC", "FRISBY BIC"],
    "HOMECENTER": ["HOMECENTER", "SODIMAC", "SODIMAC COLOMBIA", "SODIMAC COLOMBIA SA"],
    "JUMBO": ["JUMBO", "CENCOSUD", "CENCOSUD COLOMBIA", "CENCOSUD COLOMBIA S.A"],
    "OLIMPICA": ["OLIMPICA", "OLIMPICA SA", "OLIMPICA S.A"]
}

def normalizar_texto(texto):
    if not texto: return ""
    texto_norm = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto_norm.upper().strip()

def resolver_plantilla_por_emisor(emisor_nombre, plantillas_dir):
    emisor_norm = normalizar_texto(emisor_nombre)
    
    # 1. Búsqueda por diccionario de alias (incluye coincidencias parciales)
    for nombre_plantilla, alias_list in PLANTILLA_ALIASES.items():
        for alias in alias_list:
            if normalizar_texto(alias) in emisor_norm:
                path = os.path.join(plantillas_dir, f"{nombre_plantilla}.docx")
                if os.path.exists(path):
                    return path

    # 2. Búsqueda flexible en los nombres de archivos físicos .docx
    if os.path.exists(plantillas_dir):
        for filename in os.listdir(plantillas_dir):
            if filename.endswith('.docx') and filename.upper() != 'DEFECTO.DOCX':
                name_without_ext = normalizar_texto(filename[:-5])
                if name_without_ext and name_without_ext in emisor_norm:
                    path = os.path.join(plantillas_dir, filename)
                    if os.path.exists(path):
                        return path
                        
    # 3. Fallback a la plantilla por defecto si existe
    defecto_path = os.path.join(plantillas_dir, 'defecto.docx')
    if os.path.exists(defecto_path):
        return defecto_path
        
    return None

def descargar_word_factura(request, cufe):
    try:
        from django.db.models import F
        from . import word_generator
        factura = Factura.objects.get(cufe=cufe)
        
        # Registrar descarga para auditoría
        Factura.objects.filter(pk=factura.pk).update(descargas=F('descargas') + 1)
        factura.recalcular_totales()
        factura.generar_qr()
        
        file_stream = word_generator.generar_factura_word_bytes(factura)
        
        response = HttpResponse(
            file_stream.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        current_serie = factura.serie if factura.serie else "I0"
        current_folio = factura.folio if factura.folio else "00"
        response['Content-Disposition'] = f'attachment; filename="Factura_{current_serie}_{current_folio}.docx"'
        return response
    except Factura.DoesNotExist:
        return HttpResponse("Factura no encontrada.", status=404)
    except Exception as e:
        import traceback
        error_msg = f"Error generando Word: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_msg)
        return HttpResponse(error_msg, status=500, content_type="text/plain")

def previsualizar_tirilla(request, cufe):
    from django.shortcuts import get_object_or_404
    from .word_generator import generar_documento_factura
    factura = get_object_or_404(Factura, cufe=cufe)
    if not factura.qr_code:
        factura.generar_qr()

    doc = generar_documento_factura(factura)
    plantilla = factura.plantilla or 'JUMBO'
    qr_rid = 'rId6' if plantilla == 'EXITO' else 'rId8'
    bc_rid = 'rId7' if plantilla == 'EXITO' else 'rId9'

    barcode_img_src = None
    if bc_rid in doc.part.rels:
        import base64
        try:
            img_blob = doc.part.rels[bc_rid].target_part.blob
            b64 = base64.b64encode(img_blob).decode('utf-8')
            barcode_img_src = f"data:image/png;base64,{b64}"
        except Exception:
            barcode_img_src = None

    lineas_tirilla = []
    for p in doc.paragraphs:
        align = 'center' if p.alignment == 1 else ('right' if p.alignment == 2 else 'left')
        has_qr = any(qr_rid in r._element.xml for r in p.runs)
        has_bc = any(bc_rid in r._element.xml for r in p.runs)

        runs_data = []
        for r in p.runs:
            if not ('graphic' in r._element.xml):
                sz = None
                if r.font and r.font.size:
                    sz = r.font.size.pt
                runs_data.append({
                    'text': r.text,
                    'bold': bool(r.bold),
                    'size': sz
                })

        lineas_tirilla.append({
            'align': align,
            'has_qr': has_qr,
            'has_bc': has_bc,
            'is_empty': not p.text.strip() and not has_qr and not has_bc,
            'runs': runs_data,
            'text': p.text
        })

    return render(request, 'admin/facturas/tirilla_preview.html', {
        'factura': factura,
        'lineas_tirilla': lineas_tirilla,
        'barcode_img_src': barcode_img_src or getattr(factura, 'barcode_image', None),
    })
def descargar_lote_word_unificado(request, lote_id):
    try:
        from .models import LoteGeneracionFacturas
        from . import word_generator
        lote = LoteGeneracionFacturas.objects.get(pk=lote_id)
        
        stream = word_generator.generar_lote_word_unificado_bytes(lote)
        serie = lote.factura_base.serie if lote.factura_base else "VPFE"
        response = HttpResponse(
            stream.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'attachment; filename="Lote_{lote.id}_Facturas_Unificadas_{serie}.docx"'
        return response
    except LoteGeneracionFacturas.DoesNotExist:
        return HttpResponse("Lote no encontrado.", status=404)
    except Exception as e:
        import traceback
        return HttpResponse(f"Error generando Word unificado: {str(e)}\n\n{traceback.format_exc()}", status=500, content_type="text/plain")

def descargar_lote_zip(request, lote_id):
    try:
        from .models import LoteGeneracionFacturas
        from . import word_generator
        lote = LoteGeneracionFacturas.objects.get(pk=lote_id)
        
        zip_stream = word_generator.generar_lote_zip_bytes(lote)
        serie = lote.factura_base.serie if lote.factura_base else "VPFE"
        response = HttpResponse(
            zip_stream.getvalue(),
            content_type='application/zip'
        )
        response['Content-Disposition'] = f'attachment; filename="Lote_{lote.id}_Facturas_{serie}.zip"'
        return response
    except LoteGeneracionFacturas.DoesNotExist:
        return HttpResponse("Lote no encontrado.", status=404)
    except Exception as e:
        import traceback
        return HttpResponse(f"Error generando ZIP: {str(e)}\n\n{traceback.format_exc()}", status=500, content_type="text/plain")

def crear_admin(request):
    try:
        User.objects.filter(username='admin').delete()

        user = User.objects.create_superuser(
            username='admin',
            email='admin@test.com',
            password='admin123'
        )

        return HttpResponse(f"Admin recreado: {user.username}")

    except Exception as e:
        return HttpResponse(f"Error creando admin: {str(e)}", status=500)

def buscar_datos_emisor(request):
    emisor_nombre = request.GET.get('emisor_nombre', '')
    if emisor_nombre:
        # Obtain the most recent Factura with roughly the same emisor
        factura = Factura.objects.filter(emisor_nombre__icontains=emisor_nombre).order_by('-id').first()
        if factura:
            data = {
                'emisor_nit': factura.emisor_nit,
                'emisor_tipo_contribuyente': factura.emisor_tipo_contribuyente,
                'emisor_regimen_fiscal': factura.emisor_regimen_fiscal,
                'emisor_responsabilidad_tributaria': factura.emisor_responsabilidad_tributaria,
                'emisor_pais': factura.emisor_pais,
                'emisor_departamento': factura.emisor_departamento,
                'emisor_municipio': factura.emisor_municipio,
                'emisor_direccion': factura.emisor_direccion,
                'emisor_telefono': factura.emisor_telefono,
                'emisor_correo': factura.emisor_correo,
            }
            return JsonResponse(data)
    return JsonResponse({})