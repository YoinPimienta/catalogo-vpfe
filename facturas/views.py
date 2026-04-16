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
        return render(request, 'factura_detalle.html', {'factura': factura})
    except Factura.DoesNotExist:
        return render(request, 'buscar.html', {'error': "Documento no encontrado."})

def descargar_pdf_factura(request, cufe):
    try:
        from weasyprint import HTML
        factura = Factura.objects.get(cufe=cufe)
        factura.recalcular_totales()
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

def descargar_word_factura(request, cufe):
    try:
        factura = Factura.objects.get(cufe=cufe)
        factura.recalcular_totales()
        
        # 1. Buscar plantilla local física
        plantillas_dir = os.path.join(settings.BASE_DIR, 'plantillas_base')
        file_path = None
        
        if os.path.exists(plantillas_dir):
            emisor_upper = factura.emisor_nombre.upper()
            # Buscar si alguna plantilla contiene el nombre del emisor
            for filename in os.listdir(plantillas_dir):
                if filename.endswith('.docx'):
                    name_without_ext = filename[:-5].upper()
                    if name_without_ext and name_without_ext in emisor_upper and name_without_ext != 'DEFECTO':
                        file_path = os.path.join(plantillas_dir, filename)
                        break
            
            # Reintento con plantilla 'defecto.docx' si no se encontró específica
            if not file_path:
                defecto_path = os.path.join(plantillas_dir, 'defecto.docx')
                if os.path.exists(defecto_path):
                    file_path = defecto_path

        # Validación ESTRICTA física
        if not file_path or not os.path.exists(file_path):
            error_details = (
                f"Error Crítico: El archivo físico de la plantilla no existe en el servidor para el emisor '{factura.emisor_nombre}'.\n"
                f"Directorio de búsqueda: {plantillas_dir}\n"
                f"Por favor asegúrate de crear la carpeta 'plantillas_base' en la raíz de tu proyecto y colocar allí los "
                f"archivos .docx con el nombre de la empresa (Ej. 'FRISBY.docx') o uno llamado 'defecto.docx' como respaldo."
            )
            print(error_details)
            return HttpResponse(error_details, status=404, content_type="text/plain")
            
        doc = DocxTemplate(file_path)
        
        # 2. Manejar QR Base64 a Image
        qr_image = None
        if factura.qr_code and factura.qr_code.startswith('data:image/png;base64,'):
            qr_b64 = factura.qr_code.split(',')[1]
            qr_bytes = base64.b64decode(qr_b64)
            img_stream = io.BytesIO(qr_bytes)
            # 30 mm = 3 cm de ancho para el QR, ideal para 80mm
            qr_image = InlineImage(doc, img_stream, width=Mm(30))
        
        # Función auxiliar para formatear los números (si no tiene decimales lo deja como int, sino como float)
        def no_zeros(val):
            if val is None:
                return 0
            float_val = float(val)
            return int(float_val) if float_val.is_integer() else float_val

        # 3. Preparar lista de productos
        productos = []
        for p in factura.productos.all():
            productos.append({
                'nombre': p.nombre,
                'cantidad': no_zeros(p.cantidad),
                'valor_unitario': no_zeros(p.valor_unitario),
                'precio_unitario_venta': no_zeros(p.precio_unitario_venta),
                'descuento': no_zeros(p.descuento),
                'total_linea': no_zeros(p.total_linea)
            })

        # 4. Contexto de variables (placeholders)
        context = {
            'cufe': factura.cufe,
            'serie': factura.serie,
            'folio': factura.folio,
            'fecha_emision': factura.fecha_emision.strftime('%d/%m/%Y') if factura.fecha_emision else '',
            'fecha_vencimiento': factura.fecha_vencimiento.strftime('%d/%m/%Y') if factura.fecha_vencimiento else '',
            
            'forma_pago': factura.forma_pago,
            'medio_pago': factura.medio_pago,
            
            'rango_desde': factura.rango_desde,
            'rango_hasta': factura.rango_hasta,
            'numero_autorizacion': factura.numero_autorizacion,
            'vigencia': factura.vigencia,
            'xml_generado_por': factura.xml_generado_por,

            'emisor_nombre': factura.emisor_nombre,
            'emisor_nit': factura.emisor_nit,
            'emisor_direccion': factura.emisor_direccion,
            'emisor_telefono': factura.emisor_telefono,
            'emisor_correo': factura.emisor_correo,
            
            'receptor_nombre': factura.receptor_nombre,
            'receptor_nit': factura.receptor_nit,
            'receptor_direccion': factura.receptor_direccion,
            'receptor_telefono': factura.receptor_telefono,
            'receptor_correo': factura.receptor_correo,
            
            'subtotal': no_zeros(factura.subtotal),
            'descuento_total': no_zeros(factura.descuento_detalle_total),
            'total_impuesto': no_zeros(factura.total_impuesto),
            'iva': no_zeros(factura.iva),
            'inc_total': no_zeros(factura.inc_total),
            'total': no_zeros(factura.total),
            
            'notas_finales': factura.notas_finales,
            
            'productos': productos,
            'qr': qr_image if qr_image else ''
        }

        # 5. Renderizar y retornar
        doc.render(context)
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        response = HttpResponse(
            file_stream.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'attachment; filename="Factura_{factura.serie}_{factura.folio}.docx"'
        return response

    except Factura.DoesNotExist:
        return HttpResponse("Factura no encontrada.", status=404)
    except Exception as e:
        import traceback
        error_msg = f"Error generando Word: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_msg)
        return HttpResponse(error_msg, status=500, content_type="text/plain")

from django.contrib.auth.models import User

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