from django.db import migrations
from decimal import Decimal
import datetime

def restaurar_datos(apps, schema_editor):
    Factura = apps.get_model('facturas', 'Factura')
    LoteGeneracionFacturas = apps.get_model('facturas', 'LoteGeneracionFacturas')
    Producto = apps.get_model('facturas', 'Producto')
    User = apps.get_model('auth', 'User')

    admin_user = User.objects.filter(is_superuser=True).first() or User.objects.first()

    # 1. Base Jumbo
    f_jumbo = Factura.objects.filter(plantilla='JUMBO', lote_generacion__isnull=True).order_by('id').first()
    if not f_jumbo:
        f_jumbo = Factura.objects.create(
            plantilla='JUMBO',
            serie="AU06",
            folio="3425557",
            emisor_nombre='CENCOSUD COLOMBIA S.A',
            emisor_nit='900.155.107-1',
            emisor_municipio='JUMBO SANTA FE - MEDELLIN',
            emisor_direccion='AV EL POBLADO CRA 43A Tel: 6045013012',
            cajero='222222222',
            terminal='1034',
            registro='05',
            transaccion='9983',
            rem='REM09596',
            fecha_emision=datetime.date.today(),
            hora_emision=datetime.time(10, 21),
            subtotal=Decimal('35357.00'),
            iva=Decimal('8293.00'),
            total=Decimal('43650.00'),
            estado='Aceptada',
            notas_finales='GRAN CONTRIBUYENTE SEGÚN RESOLUCION DIAN No.000200 DE DICIEMBRE 27 DE 2024'
        )
        Producto.objects.create(
            factura=f_jumbo,
            codigo='3669368',
            nombre='Jugo hit surtido pg 10 lev.12 200ml c-u',
            cantidad=3,
            valor_unitario=Decimal('14550.00'),
            total_linea=Decimal('43650.00')
        )

    # 2. Base Exito
    f_exito = Factura.objects.filter(plantilla='EXITO', lote_generacion__isnull=True).order_by('id').first()
    if not f_exito:
        f_exito = Factura.objects.create(
            plantilla='EXITO',
            serie="WJ53",
            folio="44660",
            emisor_nombre='ALMACENES EXITO S.A',
            emisor_nit='8909006089',
            emisor_municipio='EXITO WOW RIONEGRO',
            emisor_telefono='3054829046',
            receptor_nombre='Dany Arturo Quintero Rojas',
            receptor_nit='7186376',
            cajero='849',
            terminal='0042',
            registro='029',
            transaccion='0197',
            rem='WJ5344660',
            fecha_emision=datetime.date.today(),
            hora_emision=datetime.time(8, 30, 8),
            subtotal=Decimal('8826220.00'),
            iva=Decimal('1676982.00'),
            total=Decimal('10503202.00'),
            estado='Aceptada',
            notas_finales='RESPONSABLE DE IVA-GRAN CONTRIBUYENTE - AUTORETENEDOR RES.8825 DE 16/NOV/2016'
        )
        Producto.objects.create(
            factura=f_exito,
            codigo='3100827',
            nombre='Nevera DOOR COOLING LG No Frost Tipo europeo 420 Lts GB41WPPAP Gris',
            cantidad=1,
            valor_unitario=Decimal('3365000.00'),
            total_linea=Decimal('3365000.00')
        )

    # 3. Base Frisby
    f_frisby = Factura.objects.filter(emisor_nombre__icontains='FRISBY', lote_generacion__isnull=True).order_by('id').first()
    if not f_frisby:
        f_frisby = Factura.objects.create(
            plantilla='OTRA',
            serie='I041',
            folio='15129088',
            emisor_nombre='FRISBY S.A. BIC',
            emisor_nit='891408584',
            emisor_direccion='CALLE 34 # 12-45',
            receptor_nombre='CLIENTE GENERAL',
            receptor_nit='222222222222',
            subtotal=Decimal('36681.00'),
            iva=Decimal('6969.00'),
            total=Decimal('43650.00'),
            fecha_emision=datetime.date.today(),
            hora_emision=datetime.time(12, 0)
        )
        Producto.objects.create(
            factura=f_frisby,
            nombre='COMBO FRISBY TRADICIONAL',
            cantidad=1,
            valor_unitario=Decimal('43650.00'),
            total_linea=Decimal('43650.00')
        )

    base_ids = [b.id for b in [f_jumbo, f_exito, f_frisby] if b]

    # 4. Eliminar facturas huerfanas duplicadas (ID no base que tengan lote_generacion NULL)
    Factura.objects.filter(lote_generacion__isnull=True).exclude(id__in=base_ids).delete()

    # 5. Asegurar existencia de los Lotes de Generacion
    if LoteGeneracionFacturas.objects.count() == 0:
        lote1 = LoteGeneracionFacturas.objects.create(
            plantilla='JUMBO',
            usuario_creador=admin_user,
            factura_base=f_jumbo,
            serie='AU06',
            folio_inicial='3425558',
            salto_folio='1',
            cantidad_generada=3,
            fecha_inicio=datetime.date.today(),
            facturas_por_dia=3,
            estado='ACTIVO'
        )
        for i in range(3):
            fol = str(3425558 + i)
            Factura.objects.create(
                plantilla='JUMBO',
                lote_generacion=lote1,
                serie='AU06',
                folio=fol,
                emisor_nombre=f_jumbo.emisor_nombre,
                emisor_nit=f_jumbo.emisor_nit,
                emisor_municipio=f_jumbo.emisor_municipio,
                emisor_direccion=f_jumbo.emisor_direccion,
                cajero=f_jumbo.cajero,
                terminal=f_jumbo.terminal,
                registro=f_jumbo.registro,
                transaccion=str(9984 + i),
                rem=f"REM{9597 + i}",
                fecha_emision=datetime.date.today(),
                hora_emision=datetime.time(10 + i, 15),
                subtotal=f_jumbo.subtotal,
                iva=f_jumbo.iva,
                total=f_jumbo.total,
                estado='Aceptada'
            )

        lote2 = LoteGeneracionFacturas.objects.create(
            plantilla='EXITO',
            usuario_creador=admin_user,
            factura_base=f_exito,
            serie='WJ53',
            folio_inicial='44661',
            salto_folio='1',
            cantidad_generada=3,
            fecha_inicio=datetime.date.today(),
            facturas_por_dia=3,
            estado='ACTIVO'
        )
        for i in range(3):
            fol = str(44661 + i)
            Factura.objects.create(
                plantilla='EXITO',
                lote_generacion=lote2,
                serie='WJ53',
                folio=fol,
                emisor_nombre=f_exito.emisor_nombre,
                emisor_nit=f_exito.emisor_nit,
                emisor_municipio=f_exito.emisor_municipio,
                emisor_telefono=f_exito.emisor_telefono,
                receptor_nombre=f_exito.receptor_nombre,
                receptor_nit=f_exito.receptor_nit,
                cajero=f_exito.cajero,
                terminal=f_exito.terminal,
                registro=f_exito.registro,
                transaccion=f"{198 + i:04d}",
                rem=f"WJ53{fol}",
                fecha_emision=datetime.date.today(),
                hora_emision=datetime.time(9 + i, 30),
                subtotal=f_exito.subtotal,
                iva=f_exito.iva,
                total=f_exito.total,
                estado='Aceptada'
            )

    try:
        from facturas.models import resetear_autoincrement_tabla
        resetear_autoincrement_tabla('facturas_factura')
        resetear_autoincrement_tabla('facturas_lotegeneracionfacturas')
    except Exception:
        pass

def revertir_migracion(apps, schema_editor):
    pass

class Migration(migrations.Migration):

    dependencies = [
        ('facturas', '0018_alter_factura_lote_generacion'),
    ]

    operations = [
        migrations.RunPython(restaurar_datos, revertir_migracion),
    ]
