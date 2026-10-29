from django.db import migrations

def corregir_plantillas(apps, schema_editor):
    Factura = apps.get_model('facturas', 'Factura')
    for f in Factura.objects.all():
        em = (f.emisor_nombre or '').upper()
        nit = (f.emisor_nit or '').replace('.', '').replace('-', '').strip()
        if 'EXITO' in em or 'ÉXITO' in em or '890900608' in nit:
            if f.plantilla != 'EXITO':
                f.plantilla = 'EXITO'
                f.save(update_fields=['plantilla'])
        elif 'CENCOSUD' in em or 'JUMBO' in em or '900155107' in nit:
            if f.plantilla != 'JUMBO':
                f.plantilla = 'JUMBO'
                f.save(update_fields=['plantilla'])

def reverse_corregir(apps, schema_editor):
    pass

class Migration(migrations.Migration):

    dependencies = [
        ('facturas', '0019_restaurar_lotes_y_limpiar_facturas'),
    ]

    operations = [
        migrations.RunPython(corregir_plantillas, reverse_corregir),
    ]
