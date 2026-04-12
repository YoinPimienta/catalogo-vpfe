from django.db import models
from datetime import date
import uuid
import qrcode
from io import BytesIO
from django.core.files import File
import base64

class Factura(models.Model):
    ESTADOS = [
        ('Aceptada', 'Aceptada'),
        ('Rechazada', 'Rechazada'),
        ('Pendiente', 'Pendiente'),
    ]

    cufe = models.CharField(max_length=150, unique=True, blank=True, help_text="Se autogenerará si se deja en blanco.")
    serie = models.CharField(max_length=20)
    folio = models.CharField(max_length=20)
    fecha_emision = models.DateField(default=date.today)
    
    # Emisor
    emisor_nombre = models.CharField(max_length=200)
    emisor_nit = models.CharField(max_length=50)
    
    # Receptor
    receptor_nombre = models.CharField(max_length=200)
    receptor_nit = models.CharField(max_length=50)
    
    # Totales
    subtotal = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    iva = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    estado = models.CharField(max_length=100, choices=ESTADOS, default='Aceptada')
    
    qr_code = models.TextField(blank=True, null=True, help_text="Se generar automticamente (Base64).")

    def save(self, *args, **kwargs):
        if not self.cufe:
            self.cufe = str(uuid.uuid4()).replace('-', '') + str(uuid.uuid4()).replace('-', '')

        super().save(*args, **kwargs)

        # Generar QR code si no tiene o la url cambia
        if not self.qr_code:
            self.generar_qr()

    def generar_qr(self):
        # La URL que el QR debe tener para redirigir a la página de búsqueda con el CUFE
        url = f"https://catalogo-vpfe.onrender.com/?cufe={self.cufe}"  
        
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        
        # Save as Base64 text!
        qr_b64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
        self.qr_code = f"data:image/png;base64,{qr_b64}"
        self.save(update_fields=['qr_code'])

    def __str__(self):
        return f"{self.serie}{self.folio} - {self.emisor_nombre}"

class Producto(models.Model):
    factura = models.ForeignKey(Factura, related_name='productos', on_delete=models.CASCADE)
    nombre = models.CharField(max_length=200)
    cantidad = models.PositiveIntegerField(default=1)
    valor_unitario = models.DecimalField(max_digits=15, decimal_places=2)
    impuesto = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total_linea = models.DecimalField(max_digits=15, decimal_places=2, blank=True)

    def save(self, *args, **kwargs):
        # Auto-calcular total de la linea
        if not self.total_linea:
            self.total_linea = (self.valor_unitario * self.cantidad) + self.impuesto
        super().save(*args, **kwargs)
        
        # Actualizar totales de factura
        # Esto sucede después de guardar el producto
        self.actualizar_totales_factura()

    def delete(self, *args, **kwargs):
        factura = self.factura
        super().delete(*args, **kwargs)
        # Actualizar totales despues de borrar
        factura.productos.all().first()
        self.actualizar_totales_factura(factura_obj=factura)

    def actualizar_totales_factura(self, factura_obj=None):
        factura = factura_obj or self.factura
        productos = factura.productos.all()
        subtotal = sum((p.valor_unitario * p.cantidad) for p in productos)
        iva = sum(p.impuesto for p in productos)
        total = subtotal + iva
        
        Factura.objects.filter(pk=factura.pk).update(subtotal=subtotal, iva=iva, total=total)

    def __str__(self):
        return self.nombre
