from django.shortcuts import render
from django.template.loader import render_to_string
from django.http import HttpResponse
from .models import Factura

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
        html_string = render_to_string('factura_pdf.html', {'factura': factura, 'request': request})
        
        # Generar el PDF
        pdf_file = HTML(string=html_string, base_url=request.build_absolute_uri()).write_pdf()
        
        current_serie = factura.serie if factura.serie else "I0"
        current_folio = factura.folio if factura.folio else "00"
        
        response = HttpResponse(pdf_file, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="Factura_{current_serie}_{current_folio}.pdf"'
        return response
    except Factura.DoesNotExist:
        return HttpResponse("Factura no encontrada.", status=404)

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