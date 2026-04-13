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

    # Identificación
    cufe = models.CharField(max_length=150, unique=True, blank=True, help_text="Se autogenerará si se deja en blanco.")
    serie = models.CharField(max_length=20, default="I041")
    folio = models.CharField(max_length=20, default="15129088")
    fecha_emision = models.DateField(default=date.today)
    fecha_vencimiento = models.DateField(null=True, blank=True)
    tipo_operacion = models.CharField(max_length=100, default="10 - Estándar")
    forma_pago = models.CharField(max_length=50, default="Contado")
    medio_pago = models.CharField(max_length=50, default="Efectivo")
    orden_pedido = models.CharField(max_length=100, blank=True, null=True)
    fecha_orden_pedido = models.DateField(null=True, blank=True)

    # Emisor
    emisor_nombre = models.CharField(max_length=200, default="FRISBY S.A. BIC")
    emisor_nit = models.CharField(max_length=50, default="891408584")
    emisor_tipo_contribuyente = models.CharField(max_length=100, default="Persona Jurídica")
    emisor_regimen_fiscal = models.CharField(max_length=100, default="O-13;O-23")
    emisor_responsabilidad_tributaria = models.CharField(max_length=100, default="01 - IVA")
    emisor_pais = models.CharField(max_length=100, default="Colombia")
    emisor_departamento = models.CharField(max_length=100, default="No Especificado")
    emisor_municipio = models.CharField(max_length=100, default="No Especificado")
    emisor_direccion = models.CharField(max_length=255, default="No Registrada")
    emisor_telefono = models.CharField(max_length=50, default="0000000000")
    emisor_correo = models.CharField(max_length=150, default="facturacion@empresa.com")
    
    # Receptor
    receptor_nombre = models.CharField(max_length=200, default="DANI QUINTERO")
    receptor_nit = models.CharField(max_length=50, default="7108376")
    receptor_tipo_documento = models.CharField(max_length=50, default="CC")
    receptor_tipo_contribuyente = models.CharField(max_length=100, default="Persona Natural")
    receptor_regimen_fiscal = models.CharField(max_length=100, default="R-99-PN")
    receptor_responsabilidad_tributaria = models.CharField(max_length=100, default="ZZ - No aplica")
    receptor_pais = models.CharField(max_length=100, default="Colombia")
    receptor_departamento = models.CharField(max_length=100, default="No Especificado")
    receptor_municipio = models.CharField(max_length=100, default="No Especificado")
    receptor_direccion = models.CharField(max_length=255, default="No Registrada")
    receptor_telefono = models.CharField(max_length=50, default="0000000000")
    receptor_correo = models.CharField(max_length=150, default="cliente@correo.com")
    
    # Textos Legales y Notas
    notas_finales = models.TextField(blank=True, null=True, default="1.- Gran Contribuyente RESPONSABLES DE IVA | SOMOS AGENTES RETENEDORES DE IVA\nEl ejemplar que reciba la DIAN se utilizará sólo para fines de control, verificación y fiscalización.")
    texto_legal = models.TextField(blank=True, null=True)
    xml_generado_por = models.CharField(max_length=200, default="Proveedor Tecnológico")
    pdf_generado_por = models.CharField(max_length=200, default="Solución Gratuita DIAN")

    # Autorizacion DIAN
    fecha_validacion_dian = models.DateField(null=True, blank=True)
    numero_autorizacion = models.CharField(max_length=100, blank=True, null=True, default="18764094886993")
    rango_desde = models.CharField(max_length=50, blank=True, null=True, default="15081476")
    rango_hasta = models.CharField(max_length=50, blank=True, null=True, default="20000000")
    vigencia = models.CharField(max_length=100, blank=True, null=True, default="2026-12-26")

    # Totales y Valores extra (Manuales y Automáticos)
    estado = models.CharField(max_length=100, choices=ESTADOS, default='Aceptada')
    moneda = models.CharField(max_length=10, default="COP")
    subtotal = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico (Total Bruto)
    descuento_detalle_total = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico
    iva = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico
    inc_total = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico
    otros_impuestos = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Manual
    total_impuesto = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico
    total_neto = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico
    descuento_global = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Manual
    recargo_global = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Manual
    total = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Automatico
    
    anticipos = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Manual
    rete_fuente = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Manual
    rete_iva = models.DecimalField(max_digits=15, decimal_places=2, default=0) # Manual
    
    qr_code = models.TextField(blank=True, null=True, help_text="Se generar automticamente (Base64).")

    def save(self, *args, **kwargs):
        if not self.cufe:
            self.cufe = str(uuid.uuid4()).replace('-', '') + str(uuid.uuid4()).replace('-', '')

        super().save(*args, **kwargs)

        if not self.qr_code:
            self.generar_qr()

    def generar_qr(self):
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
        qr_b64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
        self.qr_code = f"data:image/png;base64,{qr_b64}"
        self.save(update_fields=['qr_code'])

    def __str__(self):
        return f"{self.serie}{self.folio} - {self.emisor_nombre}"

class Producto(models.Model):
    factura = models.ForeignKey(Factura, related_name='productos', on_delete=models.CASCADE)
    codigo = models.CharField(max_length=50, blank=True, null=True)
    nombre = models.CharField(max_length=200)
    unidad_medida = models.CharField(max_length=50, default="94")
    cantidad = models.PositiveIntegerField(default=1)
    valor_unitario = models.DecimalField(max_digits=15, decimal_places=2)
    descuento = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    impuesto = models.DecimalField(max_digits=15, decimal_places=2, default=0) # IVA
    impuesto_inc = models.DecimalField(max_digits=15, decimal_places=2, default=0)
    total_linea = models.DecimalField(max_digits=15, decimal_places=2, blank=True)

    @property
    def base_linea(self):
        return (self.valor_unitario * self.cantidad) - self.impuesto - self.impuesto_inc

    @property
    def precio_base_unitario(self):
        if self.cantidad and self.cantidad > 0:
            return self.base_linea / self.cantidad
        return 0

    @property
    def precio_unitario_venta(self):
        if self.cantidad and self.cantidad > 0:
            return (self.base_linea - self.descuento) / self.cantidad
        return 0

    @property
    def porcentaje_iva(self):
        if self.base_linea > 0 and self.impuesto > 0:
            return (self.impuesto / self.base_linea) * 100
        return 0

    @property
    def porcentaje_inc(self):
        if self.base_linea > 0 and self.impuesto_inc > 0:
            return (self.impuesto_inc / self.base_linea) * 100
        return 0

    def save(self, *args, **kwargs):
        if not self.total_linea:
            # El valor_unitario ya incluye los impuestos.
            self.total_linea = (self.valor_unitario * self.cantidad) - self.descuento
        super().save(*args, **kwargs)
        self.actualizar_totales_factura()

    def delete(self, *args, **kwargs):
        factura = self.factura
        super().delete(*args, **kwargs)
        factura.productos.all().first()
        self.actualizar_totales_factura(factura_obj=factura)

    def actualizar_totales_factura(self, factura_obj=None):
        factura = factura_obj or self.factura
        productos = factura.productos.all()
        
        # El valor total de cada producto ya tiene impuestos. 
        # La base imponible (subtotal) se descompone restando los impuestos.
        subtotal = sum(((p.valor_unitario * p.cantidad) - p.impuesto - p.impuesto_inc) for p in productos)
        descuento_detalle = sum(p.descuento for p in productos)
        iva = sum(p.impuesto for p in productos)
        inc = sum(p.impuesto_inc for p in productos)
        
        # Totales auto-calculados
        total_bruto = subtotal
        total_impuesto = iva + inc + factura.otros_impuestos
        total_neto = total_bruto - descuento_detalle + total_impuesto
        total = total_neto - factura.descuento_global + factura.recargo_global
        
        Factura.objects.filter(pk=factura.pk).update(
            subtotal=total_bruto,
            descuento_detalle_total=descuento_detalle,
            iva=iva,
            inc_total=inc,
            total_impuesto=total_impuesto,
            total_neto=total_neto,
            total=total
        )

    def __str__(self):
        return self.nombre
