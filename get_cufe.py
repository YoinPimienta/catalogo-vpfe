import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'dian_sim.settings')
django.setup()

from facturas.models import Factura
f = Factura.objects.first()
if f:
    print("CUFE:", f.cufe)
else:
    print("No Factura found.")
