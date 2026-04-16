from django.contrib import admin
from .models import Factura, Producto
from django.utils.html import format_html

class ProductoInline(admin.TabularInline):
    model = Producto
    extra = 1
    fields = ('codigo', 'nombre', 'unidad_medida', 'cantidad', 'valor_unitario', 'precio_incluye_impuesto', 'descuento', 'impuesto', 'impuesto_inc', 'total_linea')
    readonly_fields = ('total_linea',)

@admin.register(Factura)
class FacturaAdmin(admin.ModelAdmin):
    list_display = ('cufe_corto', 'serie', 'folio', 'emisor_nombre', 'receptor_nombre', 'total', 'estado', 'ver_codigo_qr')
    search_fields = ('cufe', 'serie', 'folio', 'emisor_nombre', 'receptor_nombre', 'emisor_nit')
    list_filter = ('estado', 'fecha_emision')
    inlines = [ProductoInline]
    readonly_fields = ('qr_code', 'qr_code_image', 'subtotal', 'descuento_detalle_total', 'iva', 'inc_total', 'total_impuesto', 'total_neto', 'total', 'descargar_word')
    
    class Media:
        js = ('admin/js/auto_fill_emisor.js',)
    
    fieldsets = (
        ('Identificación del Documento', {
            'fields': ('cufe', 'descargar_word', 'estado', 'serie', 'folio', 'fecha_emision', 'fecha_vencimiento', 'tipo_operacion', 'forma_pago', 'medio_pago', 'orden_pedido', 'fecha_orden_pedido')
        }),
        ('Datos del Emisor', {
            'fields': ('emisor_nombre', 'emisor_nit', 'emisor_tipo_contribuyente', 'emisor_regimen_fiscal', 'emisor_responsabilidad_tributaria', 'emisor_pais', 'emisor_departamento', 'emisor_municipio', 'emisor_direccion', 'emisor_telefono', 'emisor_correo')
        }),
        ('Datos del Receptor', {
            'fields': ('receptor_nombre', 'receptor_nit', 'receptor_tipo_documento', 'receptor_tipo_contribuyente', 'receptor_regimen_fiscal', 'receptor_responsabilidad_tributaria', 'receptor_pais', 'receptor_departamento', 'receptor_municipio', 'receptor_direccion', 'receptor_telefono', 'receptor_correo')
        }),
        ('Totales Automáticos', {
            'fields': ('subtotal', 'descuento_detalle_total', 'iva', 'inc_total', 'total_impuesto', 'total_neto', 'total')
        }),
        ('Valores Finales / Manuales (Afectan el Total)', {
            'fields': ('descuento_global', 'recargo_global', 'otros_impuestos', 'anticipos', 'rete_fuente', 'rete_iva', 'moneda')
        }),
        ('Textos Legales e Información Adicional', {
            'fields': ('notas_finales', 'texto_legal', 'xml_generado_por', 'pdf_generado_por')
        }),
        ('Autorización DIAN', {
            'fields': ('fecha_validacion_dian', 'numero_autorizacion', 'rango_desde', 'rango_hasta', 'vigencia')
        }),
        ('Código QR', {
            'fields': ('qr_code', 'qr_code_image')
        }),
    )

    def cufe_corto(self, obj):
        return f"{obj.cufe[:15]}..." if obj.cufe else "-"
    cufe_corto.short_description = 'CUFE'

    def ver_codigo_qr(self, obj):
        if obj.qr_code:
            return format_html('<a href="{}" target="_blank">Ver QR</a>', obj.qr_code)
        return "-"
    ver_codigo_qr.short_description = 'QR'

    def qr_code_image(self, obj):
        if obj.qr_code:
            return format_html('<img src="{}" style="max-height: 150px;"/>', obj.qr_code)
        return "-"
    qr_code_image.short_description = 'Vista Prev. QR'

    def descargar_word(self, obj):
        if obj.pk and obj.cufe:
            from django.urls import reverse
            url = reverse('descargar_word_factura', args=[obj.cufe])
            return format_html(
                '<a class="button" href="{}" style="background-color: #417690; color: white; padding: 10px 15px; border-radius: 4px; text-decoration: none; font-weight: bold;">📥 Descargar Word (80mm)</a>',
                url
            )
        return "Guarde la factura primero para generar el Word"
    descargar_word.short_description = 'Descargar Word 80mm'

