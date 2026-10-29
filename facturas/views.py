from django.shortcuts import render, redirect, get_object_or_404
from django.template.loader import render_to_string
from django.http import HttpResponse, JsonResponse
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.models import User
from .models import Factura, LoteGeneracionFacturas
import io
import os
import re
import unicodedata
import traceback

PLANTILLA_ALIASES = {
    "EXITO": ["EXITO", "ALMACENES EXITO", "ALMACENES EXITO SA", "ALMACENES EXITO S.A"],
    "FRISBY": ["FRISBY", "FRISBY SA", "FRISBY S.A", "FRISBY S.A. BIC", "FRISBY BIC"],
    "HOMECENTER": ["HOMECENTER", "SODIMAC", "SODIMAC COLOMBIA", "SODIMAC COLOMBIA SA"],
    "JUMBO": ["JUMBO", "CENCOSUD", "CENCOSUD COLOMBIA", "CENCOSUD COLOMBIA S.A"],
    "OLIMPICA": ["OLIMPICA", "OLIMPICA SA", "OLIMPICA S.A"]
}

def normalizar_texto(texto):
    if not texto:
        return ""
    texto_norm = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8')
    return texto_norm.upper().strip()

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
        
        Factura.objects.filter(pk=factura.pk).update(descargas=F('descargas') + 1)
        factura.recalcular_totales()
        factura.generar_qr()
        html_string = render_to_string('factura_pdf.html', {'factura': factura, 'request': request})
        
        pdf_file = HTML(string=html_string, base_url=request.build_absolute_uri('/')).write_pdf()
        
        safe_serie = re.sub(r'[^a-zA-Z0-9_\-]', '', str(factura.serie or 'I0'))
        safe_folio = re.sub(r'[^a-zA-Z0-9_\-]', '', str(factura.folio or '00'))
        
        response = HttpResponse(pdf_file, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="Factura_{safe_serie}_{safe_folio}.pdf"'
        return response
    except Factura.DoesNotExist:
        return HttpResponse("Factura no encontrada.", status=404)
    except Exception as e:
        error_msg = f"Error generando PDF: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_msg)
        return HttpResponse(error_msg, status=400, content_type="text/plain; charset=utf-8")

def descargar_word_factura(request, cufe):
    try:
        from django.db.models import F
        from . import word_generator
        factura = Factura.objects.get(cufe=cufe)
        
        Factura.objects.filter(pk=factura.pk).update(descargas=F('descargas') + 1)
        factura.recalcular_totales()
        factura.generar_qr()
        
        file_stream = word_generator.generar_factura_word_bytes(factura)
        
        safe_serie = re.sub(r'[^a-zA-Z0-9_\-]', '', str(factura.serie or 'I0'))
        safe_folio = re.sub(r'[^a-zA-Z0-9_\-]', '', str(factura.folio or '00'))
        
        response = HttpResponse(
            file_stream.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'attachment; filename="Factura_{safe_serie}_{safe_folio}.docx"'
        return response
    except Factura.DoesNotExist:
        return HttpResponse("Factura no encontrada.", status=404)
    except Exception as e:
        error_msg = f"Error generando Word: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_msg)
        return HttpResponse(error_msg, status=400, content_type="text/plain; charset=utf-8")

def previsualizar_tirilla(request, cufe):
    from .word_generator import generar_documento_factura
    factura = get_object_or_404(Factura, cufe=cufe)
    if not factura.qr_code:
        factura.generar_qr()

    doc = generar_documento_factura(factura)
    plantilla = (factura.plantilla or 'JUMBO').upper()
    emisor_raw = (getattr(factura, 'emisor_nombre', '') or '').upper()
    nit_raw = (getattr(factura, 'emisor_nit', '') or '').replace('.', '').replace('-', '').strip()
    if plantilla == 'JUMBO' and ('EXITO' in emisor_raw or 'ÉXITO' in emisor_raw or '890900608' in nit_raw):
        plantilla = 'EXITO'
    elif plantilla == 'EXITO' and ('CENCOSUD' in emisor_raw or 'JUMBO' in emisor_raw or '900155107' in nit_raw):
        plantilla = 'JUMBO'

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
        from . import word_generator

        lote = LoteGeneracionFacturas.objects.filter(pk=lote_id).first()
        if not lote:
            messages.error(request, f"Lote #{lote_id} no encontrado.")
            return redirect('admin:facturas_lotegeneracionfacturas_changelist')

        stream = word_generator.generar_lote_word_unificado_bytes(lote)
        
        serie = "VPFE"
        if lote.factura_base and lote.factura_base.serie:
            serie = re.sub(r'[^a-zA-Z0-9_\-]', '', str(lote.factura_base.serie))
        elif getattr(lote, 'serie', None):
            serie = re.sub(r'[^a-zA-Z0-9_\-]', '', str(lote.serie))

        filename = f"Lote_{lote.id}_Facturas_Unificadas_{serie}.docx"
        response = HttpResponse(
            stream.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response

    except Exception as e:
        tb = traceback.format_exc()
        print(f"Error generando Word unificado: {e}\n{tb}")
        # Contingencia garantizada: descargar la primera factura individual del lote
        try:
            from . import word_generator
            lote = LoteGeneracionFacturas.objects.filter(pk=lote_id).first()
            if lote:
                f_target = lote.facturas.first() or lote.factura_base
                if f_target:
                    f_stream = word_generator.generar_factura_word_bytes(f_target)
                    response = HttpResponse(
                        f_stream.getvalue(),
                        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )
                    safe_f = re.sub(r'[^a-zA-Z0-9_\-]', '', str(f_target.folio or '01'))
                    response['Content-Disposition'] = f'attachment; filename="Lote_{lote.id}_Factura_{safe_f}.docx"'
                    return response
        except Exception as e_inner:
            print(f"Error en contingencia de descarga individual: {e_inner}")

        messages.error(request, f"No se pudo generar el Word unificado para el lote #{lote_id}: {str(e)}")
        referer = request.META.get('HTTP_REFERER')
        if referer:
            return redirect(referer)
        return redirect('admin:facturas_lotegeneracionfacturas_changelist')

def descargar_lote_zip(request, lote_id):
    try:
        from . import word_generator

        lote = LoteGeneracionFacturas.objects.filter(pk=lote_id).first()
        if not lote:
            messages.error(request, f"Lote #{lote_id} no encontrado.")
            return redirect('admin:facturas_lotegeneracionfacturas_changelist')

        zip_stream = word_generator.generar_lote_zip_bytes(lote)
        
        serie = "VPFE"
        if lote.factura_base and lote.factura_base.serie:
            serie = re.sub(r'[^a-zA-Z0-9_\-]', '', str(lote.factura_base.serie))
        elif getattr(lote, 'serie', None):
            serie = re.sub(r'[^a-zA-Z0-9_\-]', '', str(lote.serie))

        filename = f"Lote_{lote.id}_Facturas_{serie}.zip"
        response = HttpResponse(
            zip_stream.getvalue(),
            content_type='application/zip'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response

    except Exception as e:
        tb = traceback.format_exc()
        print(f"Error generando ZIP lote {lote_id}: {e}\n{tb}")
        messages.error(request, f"No se pudo generar el ZIP para el lote #{lote_id}: {str(e)}")
        referer = request.META.get('HTTP_REFERER')
        if referer:
            return redirect(referer)
        return redirect('admin:facturas_lotegeneracionfacturas_changelist')

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
        factura = Factura.objects.filter(emisor_nombre__icontains=emisor_nombre).order_by('-id').first()
        if factura:
            data = {
                'plantilla': factura.plantilla,
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
