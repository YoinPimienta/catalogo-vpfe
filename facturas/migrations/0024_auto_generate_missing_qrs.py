from django.db import migrations
import uuid
import qrcode
import io
import base64

def poblar_qrs_faltantes(apps, schema_editor):
    Factura = apps.get_model('facturas', 'Factura')
    for f in Factura.objects.all():
        needs_save = False
        if not f.cufe:
            f.cufe = str(uuid.uuid4()).replace('-', '') + str(uuid.uuid4()).replace('-', '')
            needs_save = True
        if not f.qr_code:
            url = f"https://catalogo-vpfe1.onrender.com/?cufe={f.cufe}"
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(url)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            buffer = io.BytesIO()
            try:
                getattr(img, 'save')(buffer, **{"format": "PNG"})
            except TypeError:
                img.save(buffer)
            qr_b64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
            f.qr_code = f"data:image/png;base64,{qr_b64}"
            needs_save = True
        if needs_save:
            f.save()

class Migration(migrations.Migration):

    dependencies = [
        ('facturas', '0023_alter_lotegeneracionfacturas_factura_base'),
    ]

    operations = [
        migrations.RunPython(poblar_qrs_faltantes, migrations.RunPython.noop),
    ]
