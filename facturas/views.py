from django.shortcuts import render
from .models import Factura

def buscar_factura(request):
    cufe = request.GET.get('cufe') or request.POST.get('cufe')
    error = None
    factura = None

    if cufe:
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

from django.contrib.auth.models import User

def crear_admin(request):
    if not User.objects.filter(username='admin').exists():
        User.objects.create_superuser(
            username='admin',
            email='admin@test.com',
            password='admin123'
        )
        return HttpResponse("Admin creado")
    return HttpResponse("Ya existe")