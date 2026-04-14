from django.shortcuts import render
from django.template.loader import render_to_string
from django.http import HttpResponse
from .models import Factura, PlantillaWord
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import base64

def buscar_factura(request):
    cufe = request.GET.get('cufe') or request.POST.get('cufe')
    error = None
    factura = None

    if request.method == 'POST' and cufe:
        try:
            factura = Factura.objects.get(cufe=cufe)
            return render(request, 'factura_detalle.html', {'factura': factura})
        except Factura.DoesNotExist:
            error = "Documento no encontrado. Verifique el código ingresado e intente nuevamente."

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
        
        # 1. Buscar plantilla
        plantilla = PlantillaWord.objects.filter(empresa__iexact=factura.emisor_nombre, activa=True).first()
        if not plantilla:
            plantilla = PlantillaWord.objects.filter(por_defecto=True, activa=True).first()
            
        if not plantilla or not plantilla.archivo:
            return HttpResponse("No se encontró una plantilla de Word asociada para este emisor ni una plantilla por defecto.", status=404)
            
        doc = DocxTemplate(plantilla.archivo.path)
        
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