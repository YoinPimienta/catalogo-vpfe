from django.shortcuts import render
from .models import Factura

def buscar_factura(request):
    error = None
    if request.method == 'POST':
        cufe_uuid = request.POST.get('cufe', '').strip()
        if not cufe_uuid:
            error = "CUFE o UUID es requerido."
        else:
            try:
                factura = Factura.objects.get(cufe=cufe_uuid)
                return render(request, 'factura_detalle.html', {'factura': factura})
            except Factura.DoesNotExist:
                error = "Documento no encontrado. Verifique el código ingresado e intente nuevamente."

    return render(request, 'buscar.html', {'error': error})

def ver_factura(request, cufe):
    try:
        factura = Factura.objects.get(cufe=cufe)
        return render(request, 'factura_detalle.html', {'factura': factura})
    except Factura.DoesNotExist:
        return render(request, 'buscar.html', {'error': "Documento no encontrado."})
