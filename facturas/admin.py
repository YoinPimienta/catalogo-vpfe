from django.contrib import admin
from .models import Factura, Producto
from django.utils.html import format_html

class ProductoInline(admin.TabularInline):
    model = Producto
    extra = 1

@admin.register(Factura)
class FacturaAdmin(admin.ModelAdmin):
    list_display = ('cufe_corto', 'serie', 'folio', 'emisor_nombre', 'receptor_nombre', 'total', 'estado', 'ver_codigo_qr')
    search_fields = ('cufe', 'serie', 'folio', 'emisor_nombre', 'receptor_nombre', 'emisor_nit')
    list_filter = ('estado', 'fecha_emision')
    inlines = [ProductoInline]
    readonly_fields = ('qr_code_image', 'subtotal', 'iva', 'total')
    
    fieldsets = (
        ('Identificación del Documento', {
            'fields': (('cufe', 'estado'), ('serie', 'folio', 'fecha_emision'))
        }),
        ('Datos del Emisor', {
            'fields': (('emisor_nombre', 'emisor_nit'),)
        }),
        ('Datos del Receptor', {
            'fields': (('receptor_nombre', 'receptor_nit'),)
        }),
        ('Totales', {
            'fields': (('subtotal', 'iva', 'total'),)
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
            return format_html('<a href="{}" target="_blank">Ver QR</a>', obj.qr_code.url)
        return "-"
    ver_codigo_qr.short_description = 'QR'

    def qr_code_image(self, obj):
        if obj.qr_code:
            return format_html('<img src="{}" style="max-height: 150px;"/>', obj.qr_code.url)
        return "-"
    qr_code_image.short_description = 'Vista Prev. QR'
