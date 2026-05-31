from django.contrib import admin
from .models import Factura, Producto, LoteGeneracionFacturas
from django.utils.html import format_html
from django import forms
from django.shortcuts import get_object_or_404, render, redirect
from django.urls import path, reverse
from django.contrib import messages
from django.db import transaction
from django.utils import timezone
from collections import Counter
import datetime
import re

# Formulario administrativo para configurar la generación masiva
class GenerarFacturasMasivasForm(forms.Form):
    cantidad = forms.IntegerField(
        min_value=1,
        max_value=500,
        initial=10,
        label="Cantidad Total de Facturas",
        help_text="Número total de facturas copias a generar (Max: 500)."
    )
    fecha_inicial = forms.DateField(
        initial=datetime.date.today,
        label="Fecha Inicial",
        widget=forms.DateInput(attrs={'type': 'date', 'style': 'max-width: 250px;'}),
        help_text="Fecha de emisión del primer documento a generar."
    )
    fecha_final = forms.DateField(
        required=False,
        label="Fecha Final (Límite)",
        widget=forms.DateInput(attrs={'type': 'date', 'style': 'max-width: 250px;'}),
        help_text="Opcional. Límite de fecha. Si la secuencia la supera, la generación se detendrá preventivamente."
    )
    facturas_por_dia = forms.IntegerField(
        min_value=1,
        initial=1,
        label="Facturas por Día",
        help_text="Cuántas facturas se emitirán por cada día calendario."
    )
    folio_inicial = forms.CharField(
        max_length=20,
        required=False,
        label="Folio Inicial",
        widget=forms.TextInput(attrs={'placeholder': 'Ej: 1001', 'style': 'max-width: 250px;'}),
        help_text="Si se deja vacío, se tomará el de la factura base incrementado en 1."
    )
    incremento_folio = forms.IntegerField(
        min_value=1,
        initial=1,
        label="Incrementar Folio en",
        help_text="Valor en el que se incrementará consecutivamente el folio (por defecto 1)."
    )
    comportamiento_vencimiento = forms.ChoiceField(
        choices=[
            ('plazo_relativo', 'Mantener plazo relativo (plazo en días de la factura base)'),
            ('misma_fecha', 'Misma fecha de emisión'),
            ('sin_vencimiento', 'Sin fecha de vencimiento (Vacío)'),
        ],
        initial='plazo_relativo',
        label="Comportamiento del Vencimiento",
        help_text="Cómo calcular la fecha de vencimiento respecto a la de emisión."
    )

# Función auxiliar robusta para incrementar folios alfanuméricos manteniendo ceros a la izquierda
def increment_folio(folio_str, increment=1):
    match = re.search(r'(\d+)$', folio_str)
    if match:
        digits = match.group(1)
        num = int(digits) + increment
        new_digits = f"{num:0{len(digits)}d}"
        return folio_str[:match.start()] + new_digits
    else:
        return f"{folio_str}_{increment}"

# Comparador estricto para detectar si una factura generada sufrió modificaciones manuales respecto a la base
def tiene_cambios_respecto_a_base(factura, factura_base):
    # 1. Comparar total, subtotal e información de receptor (redondeando a 2 decimales para evitar inconsistencias de escala)
    from decimal import Decimal
    t_fact = round(Decimal(str(factura.total)), 2)
    t_base = round(Decimal(str(factura_base.total)), 2)
    s_fact = round(Decimal(str(factura.subtotal)), 2)
    s_base = round(Decimal(str(factura_base.subtotal)), 2)

    if t_fact != t_base or s_fact != s_base:
        return True
    if factura.receptor_nit != factura_base.receptor_nit or factura.receptor_nombre != factura_base.receptor_nombre:
        return True
    
    # 2. Comparar productos y montos unitarios
    prod_factura = list(factura.productos.all().order_by('id'))
    prod_base = list(factura_base.productos.all().order_by('id'))
    
    if len(prod_factura) != len(prod_base):
        return True
        
    for pf, pb in zip(prod_factura, prod_base):
        val_pf = round(Decimal(str(pf.valor_unitario)), 2)
        val_pb = round(Decimal(str(pb.valor_unitario)), 2)
        desc_pf = round(Decimal(str(pf.descuento)), 2)
        desc_pb = round(Decimal(str(pb.descuento)), 2)
        
        if (pf.nombre != pb.nombre or 
            pf.cantidad != pb.cantidad or 
            val_pf != val_pb or 
            desc_pf != desc_pb or
            pf.impuesto != pb.impuesto or
            pf.impuesto_inc != pb.impuesto_inc):
            return True
            
    return False


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
    
    change_form_template = 'admin/facturas/factura/change_form.html'
    
    class Media:
        js = ('admin/js/auto_fill_emisor.js',)
    
    fieldsets = (
        ('Identificación del Documento', {
            'fields': ('cufe', 'descargar_word', 'estado', 'serie', 'folio', 'fecha_emision', 'fecha_vencimiento', 'tipo_operacion', 'forma_pago', 'orden_pedido', 'fecha_orden_pedido')
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

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                '<path:object_id>/generar-masivo/',
                self.admin_site.admin_view(self.generar_masivo_view),
                name='facturas_factura_generar_masivo',
            ),
        ]
        return custom_urls + urls

    # Vista de generación masiva adaptada para auditar y registrar el lote
    def generar_masivo_view(self, request, object_id):
        factura_base = get_object_or_404(Factura, pk=object_id)
        error_message = None

        if request.method == 'POST':
            form = GenerarFacturasMasivasForm(request.POST)
            if form.is_valid():
                cantidad = form.cleaned_data['cantidad']
                fecha_inicial = form.cleaned_data['fecha_inicial']
                fecha_final = form.cleaned_data['fecha_final']
                facturas_por_dia = form.cleaned_data['facturas_por_dia']
                folio_inicial = form.cleaned_data['folio_inicial']
                incremento_folio_val = form.cleaned_data['incremento_folio']
                comportamiento_vencimiento = form.cleaned_data['comportamiento_vencimiento']

                if not folio_inicial:
                    folio_inicial = increment_folio(factura_base.folio, 1)

                dias_plazo = 0
                if factura_base.fecha_emision and factura_base.fecha_vencimiento:
                    dias_plazo = (factura_base.fecha_vencimiento - factura_base.fecha_emision).days

                try:
                    with transaction.atomic():
                        facturas_creadas = []
                        folios_detectados_duplicados = []

                        # Primera pasada: Validar fechas y folios duplicados preventivamente
                        for i in range(cantidad):
                            dia_offset = i // facturas_por_dia
                            fecha_emision_nueva = fecha_inicial + datetime.timedelta(days=dia_offset)

                            if fecha_final and fecha_emision_nueva > fecha_final:
                                raise forms.ValidationError(
                                    f"La fecha de emisión calculada para la factura {i+1} ({fecha_emision_nueva}) supera la fecha límite establecida ({fecha_final}). Proceso cancelado."
                                )

                            folio_nuevo = increment_folio(folio_inicial, i * incremento_folio_val)

                            if Factura.objects.filter(serie=factura_base.serie, folio=folio_nuevo).exists():
                                folios_detectados_duplicados.append(folio_nuevo)

                        if folios_detectados_duplicados:
                            raise forms.ValidationError(
                                f"Conflicto de folios duplicados detectado. Los siguientes folios ya existen para la serie '{factura_base.serie}': {', '.join(folios_detectados_duplicados)}. Proceso abortado sin realizar cambios en la base de datos."
                            )

                        # Crear el Lote de Generación de Facturas
                        fecha_fin_calculada = fecha_inicial + datetime.timedelta(days=((cantidad - 1) // facturas_por_dia))
                        lote = LoteGeneracionFacturas.objects.create(
                            usuario_creador=request.user,
                            factura_base=factura_base,
                            cantidad_generada=cantidad,
                            fecha_inicio=fecha_inicial,
                            fecha_fin=fecha_fin_calculada,
                            facturas_por_dia=facturas_por_dia,
                            observaciones=f"Generación consecutiva de {cantidad} facturas.",
                            estado='ACTIVO'
                        )

                        # Segunda pasada: Clonar Factura y Productos
                        for i in range(cantidad):
                            dia_offset = i // facturas_por_dia
                            fecha_emision_nueva = fecha_inicial + datetime.timedelta(days=dia_offset)
                            
                            if comportamiento_vencimiento == 'plazo_relativo':
                                fecha_vencimiento_nueva = fecha_emision_nueva + datetime.timedelta(days=dias_plazo) if dias_plazo > 0 else None
                            elif comportamiento_vencimiento == 'misma_fecha':
                                fecha_vencimiento_nueva = fecha_emision_nueva
                            else:
                                fecha_vencimiento_nueva = None
                            
                            folio_nuevo = increment_folio(folio_inicial, i * incremento_folio_val)

                            copia_factura = Factura(
                                lote_generacion=lote,
                                serie=factura_base.serie,
                                folio=folio_nuevo,
                                fecha_emision=fecha_emision_nueva,
                                fecha_vencimiento=fecha_vencimiento_nueva,
                                cufe="",
                                qr_code="",
                                tipo_operacion=factura_base.tipo_operacion,
                                forma_pago=factura_base.forma_pago,
                                medio_pago=factura_base.medio_pago,
                                orden_pedido=factura_base.orden_pedido,
                                fecha_orden_pedido=factura_base.fecha_orden_pedido,
                                
                                # Emisor
                                emisor_nombre=factura_base.emisor_nombre,
                                emisor_nit=factura_base.emisor_nit,
                                emisor_tipo_contribuyente=factura_base.emisor_tipo_contribuyente,
                                emisor_regimen_fiscal=factura_base.emisor_regimen_fiscal,
                                emisor_responsabilidad_tributaria=factura_base.emisor_responsabilidad_tributaria,
                                emisor_pais=factura_base.emisor_pais,
                                emisor_departamento=factura_base.emisor_departamento,
                                emisor_municipio=factura_base.emisor_municipio,
                                emisor_direccion=factura_base.emisor_direccion,
                                emisor_telefono=factura_base.emisor_telefono,
                                emisor_correo=factura_base.emisor_correo,
                                
                                # Receptor
                                receptor_nombre=factura_base.receptor_nombre,
                                receptor_nit=factura_base.receptor_nit,
                                receptor_tipo_documento=factura_base.receptor_tipo_documento,
                                receptor_tipo_contribuyente=factura_base.receptor_tipo_contribuyente,
                                receptor_regimen_fiscal=factura_base.receptor_regimen_fiscal,
                                receptor_responsabilidad_tributaria=factura_base.receptor_responsabilidad_tributaria,
                                receptor_pais=factura_base.receptor_pais,
                                receptor_departamento=factura_base.receptor_departamento,
                                receptor_municipio=factura_base.receptor_municipio,
                                receptor_direccion=factura_base.receptor_direccion,
                                receptor_telefono=factura_base.receptor_telefono,
                                receptor_correo=factura_base.receptor_correo,
                                
                                # Textos y DIAN
                                notas_finales=getattr(factura_base, 'notas_finales', ''),
                                texto_legal=factura_base.texto_legal,
                                xml_generado_por=factura_base.xml_generado_por,
                                pdf_generado_por=factura_base.pdf_generado_por,
                                numero_autorizacion=factura_base.numero_autorizacion,
                                rango_desde=factura_base.rango_desde,
                                rango_hasta=factura_base.rango_hasta,
                                vigencia=factura_base.vigencia,
                                fecha_validacion_dian=factura_base.fecha_validacion_dian,
                                
                                estado=factura_base.estado,
                                moneda=factura_base.moneda,
                                otros_impuestos=factura_base.otros_impuestos,
                                descuento_global=factura_base.descuento_global,
                                recargo_global=factura_base.recargo_global,
                                anticipos=factura_base.anticipos,
                                rete_fuente=factura_base.rete_fuente,
                                rete_iva=factura_base.rete_iva,
                            )
                            copia_factura.save()

                            for producto_base in factura_base.productos.all():
                                Producto.objects.create(
                                    factura=copia_factura,
                                    codigo=producto_base.codigo,
                                    nombre=producto_base.nombre,
                                    unidad_medida=producto_base.unidad_medida,
                                    cantidad=producto_base.cantidad,
                                    valor_unitario=producto_base.valor_unitario,
                                    precio_incluye_impuesto=producto_base.precio_incluye_impuesto,
                                    descuento=producto_base.descuento,
                                    impuesto=producto_base.impuesto,
                                    impuesto_inc=producto_base.impuesto_inc,
                                    total_linea=producto_base.total_linea
                                )

                            copia_factura.recalcular_totales()
                            facturas_creadas.append(copia_factura)

                        folios_creados = [f"{f.serie}{f.folio}" for f in facturas_creadas]
                        rango_folios = f"desde {folios_creados[0]} hasta {folios_creados[-1]}" if len(folios_creados) > 1 else folios_creados[0]
                        messages.success(
                            request,
                            f"Se generaron con éxito {len(facturas_creadas)} facturas masivas ({rango_folios}) en el Lote #{lote.id}."
                        )
                        return redirect('admin:facturas_factura_changelist')

                except forms.ValidationError as e:
                    error_message = e.message
                except Exception as e:
                    error_message = f"Error crítico durante el procesamiento en base de datos: {str(e)}"

        else:
            siguiente_folio = increment_folio(factura_base.folio, 1)
            form = GenerarFacturasMasivasForm(initial={
                'fecha_inicial': datetime.date.today(),
                'folio_inicial': siguiente_folio,
            })

        context = {
            **self.admin_site.each_context(request),
            'opts': self.model._meta,
            'original': factura_base,
            'form': form,
            'error_message': error_message,
        }
        return render(request, 'admin/facturas/factura/generar_masivo.html', context)


@admin.register(LoteGeneracionFacturas)
class LoteGeneracionFacturasAdmin(admin.ModelAdmin):
    list_display = ('id', 'usuario_creador', 'fecha_creacion', 'factura_base_link', 'cantidad_generada', 'estado_badge')
    list_filter = ('estado', 'fecha_creacion')
    search_fields = ('id', 'usuario_creador__username', 'factura_base__serie', 'factura_base__folio')
    readonly_fields = (
        'usuario_creador', 'usuario_cancelador', 'fecha_creacion', 'fecha_cancelacion', 
        'factura_base', 'cantidad_generada', 'fecha_inicio', 'fecha_fin', 
        'facturas_por_dia', 'observaciones', 'motivo_cancelacion', 'estado', 
        'mostrar_estadisticas', 'listado_facturas'
    )
    change_form_template = 'admin/facturas/lotegeneracionfacturas/change_form.html'
    
    fieldsets = (
        ('Información de Registro', {
            'fields': ('usuario_creador', 'fecha_creacion', 'factura_base', 'estado')
        }),
        ('Detalle de Generación Masiva', {
            'fields': ('cantidad_generada', 'fecha_inicio', 'fecha_fin', 'facturas_por_dia', 'observaciones')
        }),
        ('Estadísticas de Facturación', {
            'fields': ('mostrar_estadisticas',)
        }),
        ('Trazabilidad e Historial del Lote', {
            'fields': ('listado_facturas',)
        }),
        ('Auditoría de Reversión / Anulación', {
            'fields': ('usuario_cancelador', 'fecha_cancelacion', 'motivo_cancelacion')
        }),
    )

    def factura_base_link(self, obj):
        if obj.factura_base:
            url = reverse('admin:facturas_factura_change', args=[obj.factura_base.pk])
            return format_html('<a href="{}">{}</a>', url, f"{obj.factura_base.serie}{obj.factura_base.folio}")
        return "-"
    factura_base_link.short_description = 'Factura Base'

    def estado_badge(self, obj):
        colors = {
            'ACTIVO': '#047857',      # emerald green
            'CANCELADO': '#be123c',   # rose red
            'ELIMINADO': '#4b5563',   # gray
            'DESACTIVADO': '#b45309', # amber orange
            'PARCIAL': '#4338ca',     # indigo blue
        }
        color = colors.get(obj.estado, '#000000')
        return format_html(
            '<span style="background-color: {}; color: white; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 0.75rem;">{}</span>',
            color, obj.estado
        )
    estado_badge.short_description = 'Estado'

    def mostrar_estadisticas(self, obj):
        if not obj or not obj.pk:
            return "Las estadísticas estarán disponibles tras guardar el lote."
        
        facturas = obj.facturas.all()
        total_facturas = facturas.count()
        if total_facturas == 0:
            return "Este lote no contiene facturas generadas."
        
        suma_facturada = sum(f.total for f in facturas)
        fechas = [f.fecha_emision for f in facturas if f.fecha_emision]
        fecha_ini = min(fechas) if fechas else obj.fecha_inicio
        fecha_fin = max(fechas) if fechas else obj.fecha_fin
        
        conteo_dias = Counter(fechas)
        promedio_diario = total_facturas / len(conteo_dias) if conteo_dias else 0
        
        html = f"""
        <div style="background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 20px; font-family: sans-serif; color: #1e293b;">
            <h4 style="margin-top:0; color:#1e3a47; border-bottom: 2px solid #248080; padding-bottom: 8px; font-size: 1.1rem; font-weight: bold; display: flex; align-items: center; gap: 6px;">
                📊 Resumen de Rendimiento de Facturación
            </h4>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-top: 15px;">
                <div style="background: #ffffff; padding: 12px; border-radius: 6px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
                    <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.5px; color:#64748b; display:block; margin-bottom:4px;">Total Copias</span>
                    <strong style="font-size:1.3rem; color:#248080;">{total_facturas}</strong>
                </div>
                <div style="background: #ffffff; padding: 12px; border-radius: 6px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
                    <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.5px; color:#64748b; display:block; margin-bottom:4px;">Monto Total Facturado</span>
                    <strong style="font-size:1.3rem; color:#16a34a;">${suma_facturada:,.2f} COP</strong>
                </div>
                <div style="background: #ffffff; padding: 12px; border-radius: 6px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
                    <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.5px; color:#64748b; display:block; margin-bottom:4px;">Rango de Emisión</span>
                    <strong style="font-size:0.95rem; color:#0f172a; display:block; margin-top:4px;">
                        {fecha_ini.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y') if fecha_fin else '-'}
                    </strong>
                </div>
                <div style="background: #ffffff; padding: 12px; border-radius: 6px; border: 1px solid #e2e8f0; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
                    <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.5px; color:#64748b; display:block; margin-bottom:4px;">Promedio por Día</span>
                    <strong style="font-size:1.3rem; color:#2563eb;">{promedio_diario:.1f} <span style="font-size:0.8rem; font-weight:normal; color:#475569;">doc/día</span></strong>
                </div>
            </div>
            
            <h5 style="margin-bottom:8px; margin-top:20px; color:#475569; font-weight: bold;">📅 Distribución Consecutiva por Día Calendario:</h5>
            <div style="max-height: 120px; overflow-y: auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px;">
                <ul style="margin: 0; padding-left: 20px; font-size: 0.85rem; line-height: 1.6;">
        """
        for dia, cant in sorted(conteo_dias.items()):
            html += f"<li>Fecha <strong>{dia.strftime('%d/%m/%Y')}</strong>: <span style='color:#248080; font-weight:bold;'>{cant}</span> factura(s) copias</li>"
        
        html += "</ul></div></div>"
        return format_html(html)
    mostrar_estadisticas.short_description = 'Estadísticas Consolidadas del Lote'

    def listado_facturas(self, obj):
        if not obj or not obj.pk:
            return "-"
        facturas = obj.facturas.all()
        if not facturas.exists():
            return "No se encontraron facturas asociadas a este lote."
        
        html = """
        <div style="max-height: 350px; overflow-y: auto; border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
        <table style="width:100%; border-collapse:collapse; font-size:0.85rem; text-align:left; font-family: sans-serif;">
            <thead>
                <tr style="background:#f1f5f9; border-bottom:2px solid #cbd5e1; color:#475569;">
                    <th style="padding:10px;">Folio Consecutivo</th>
                    <th style="padding:10px;">CUFE Único</th>
                    <th style="padding:10px;">Fecha Emisión</th>
                    <th style="padding:10px;">Total</th>
                    <th style="padding:10px; text-align:center;">Descargas</th>
                    <th style="padding:10px;">Estado</th>
                    <th style="padding:10px; text-align:center;">Acción</th>
                </tr>
            </thead>
            <tbody>
        """
        for f in facturas:
            url = reverse('admin:facturas_factura_change', args=[f.pk])
            descargas_warning = f'<span style="color:#dc2626; font-weight:bold;">{f.descargas}</span>' if f.descargas > 0 else '0'
            html += f"""
                <tr style="border-bottom:1px solid #f1f5f9; color:#1e293b;">
                    <td style="padding:10px; font-weight:bold;">{f.serie}{f.folio}</td>
                    <td style="padding:10px; font-family:monospace; color:#3b82f6;">{f.cufe[:24]}...</td>
                    <td style="padding:10px;">{f.fecha_emision.strftime('%d/%m/%Y') if f.fecha_emision else '-'}</td>
                    <td style="padding:10px; font-weight:bold; color:#16a34a;">${f.total:,.2f} COP</td>
                    <td style="padding:10px; text-align:center;">{descargas_warning}</td>
                    <td style="padding:10px;">{f.estado}</td>
                    <td style="padding:10px; text-align:center;">
                        <a href="{url}" target="_blank" style="background:#e2e8f0; color:#475569; padding:4px 8px; border-radius:4px; text-decoration:none; font-weight:bold; font-size:0.75rem;">👁️ Abrir</a>
                    </td>
                </tr>
            """
        html += "</tbody></table></div>"
        return format_html(html)
    listado_facturas.short_description = 'Listado e Historial de Documentos'

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('<path:object_id>/cancelar/', self.admin_site.admin_view(self.cancelar_lote_view), name='facturas_lote_cancelar'),
            path('<path:object_id>/desactivar/', self.admin_site.admin_view(self.desactivar_lote_view), name='facturas_lote_desactivar'),
        ]
        return custom_urls + urls

    def cancelar_lote_view(self, request, object_id):
        lote = get_object_or_404(LoteGeneracionFacturas, pk=object_id)
        if lote.estado != 'ACTIVO':
            messages.error(request, "Este lote ya no se encuentra en estado Activo.")
            return redirect('admin:facturas_lotegeneracionfacturas_change', lote.pk)

        facturas = lote.facturas.all()
        factura_base = lote.factura_base

        # Análisis de riesgos completo
        analisis_facturas = []
        total_con_riesgo = 0
        for f in facturas:
            tiene_cambios = False
            if factura_base:
                tiene_cambios = tiene_cambios_respecto_a_base(f, factura_base)
            
            nivel_riesgo = 'NINGUNO'
            if f.descargas > 0:
                nivel_riesgo = 'ALTO'
                total_con_riesgo += 1
            elif tiene_cambios:
                nivel_riesgo = 'MEDIO'
                total_con_riesgo += 1

            analisis_facturas.append({
                'factura': f,
                'tiene_cambios': tiene_cambios,
                'nivel_riesgo': nivel_riesgo,
                'base_total': factura_base.total if factura_base else 0
            })

        if request.method == 'POST':
            motivo = request.POST.get('motivo', '').strip()
            if not motivo:
                messages.error(request, "El motivo de la cancelación es estrictamente obligatorio para la auditoría.")
            elif total_con_riesgo > 0 and not request.user.is_superuser:
                messages.error(request, "No tienes los privilegios de superusuario requeridos para forzar la reversión de un lote con riesgos.")
            else:
                try:
                    with transaction.atomic():
                        total_eliminadas = facturas.count()
                        # La eliminación en cascada purga automáticamente los Productos de la base
                        facturas.delete()

                        # Firma de auditoría
                        lote.estado = 'CANCELADO'
                        lote.usuario_cancelador = request.user
                        lote.fecha_cancelacion = timezone.now()
                        lote.motivo_cancelacion = motivo
                        lote.save()

                        messages.success(
                            request, 
                            f"Lote #{lote.id} revertido con éxito. Se eliminaron definitivamente {total_eliminadas} facturas y sus productos asociados. Registro de auditoría guardado."
                        )
                        return redirect('admin:facturas_lotegeneracionfacturas_changelist')
                except Exception as e:
                    messages.error(request, f"Error en la reversión de base de datos: {str(e)}")

        context = {
            **self.admin_site.each_context(request),
            'opts': self.model._meta,
            'lote': lote,
            'total_facturas': facturas.count(),
            'analisis_facturas': analisis_facturas,
            'total_con_riesgo': total_con_riesgo,
            'es_superuser': request.user.is_superuser,
        }
        return render(request, 'admin/facturas/lotegeneracionfacturas/confirmar_cancelacion.html', context)

    def deactivate_lote_view(self, request, object_id):
        # Fallback support
        return self.desactivar_lote_view(request, object_id)

    def desactivar_lote_view(self, request, object_id):
        lote = get_object_or_404(LoteGeneracionFacturas, pk=object_id)
        if lote.estado != 'ACTIVO':
            messages.error(request, "Este lote ya no se encuentra en estado Activo.")
            return redirect('admin:facturas_lotegeneracionfacturas_change', lote.pk)

        facturas = lote.facturas.all()

        if request.method == 'POST':
            motivo = request.POST.get('motivo', '').strip()
            if not motivo:
                messages.error(request, "El motivo de la desactivación es estrictamente obligatorio para la auditoría.")
            else:
                try:
                    with transaction.atomic():
                        # Marcar todas las facturas asociadas como 'Anulada'
                        total_anuladas = facturas.update(estado='Anulada')

                        # Guardar firmas de auditoría
                        lote.estado = 'DESACTIVADO'
                        lote.usuario_cancelador = request.user
                        lote.fecha_cancelacion = timezone.now()
                        lote.motivo_cancelacion = motivo
                        lote.save()

                        messages.success(
                            request, 
                            f"Lote #{lote.id} desactivado en modo seguro. Las {total_anuladas} facturas se conservan intactas en la base de datos pero marcadas como 'Anuladas' y excluidas de las búsquedas normales."
                        )
                        return redirect('admin:facturas_lotegeneracionfacturas_changelist')
                except Exception as e:
                    messages.error(request, f"Error en la base de datos al desactivar el lote: {str(e)}")

        context = {
            **self.admin_site.each_context(request),
            'opts': self.model._meta,
            'lote': lote,
            'total_facturas': facturas.count(),
        }
        return render(request, 'admin/facturas/lotegeneracionfacturas/confirmar_desactivacion.html', context)
