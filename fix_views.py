import codecs

filepath = 'facturas/views.py'
with codecs.open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

new_view = """def buscar_factura(request):
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
"""

with codecs.open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines[:3])    # from django.shortcuts import render\nfrom .models import Factura\n\n
    f.write(new_view)
    f.writelines(lines[16:])   # def ver_factura(...) goes from line 17 but 0-indexed is 16
print('done')
