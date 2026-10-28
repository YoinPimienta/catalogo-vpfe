from django.contrib import admin
from .models import Factura, Producto, LoteGeneracionFacturas
from django.utils.html import format_html
from django.utils.safestring import mark_safe
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
    plantilla = forms.ChoiceField(
        choices=Factura.PLANTILLAS_CHOICES,
        initial='JUMBO',
        label="Plantilla / Formato",
        help_text="Elige entre JUMBO (Cencosud 80mm) o ÉXITO (Almacenes Éxito 80mm)."
    )
    serie = forms.CharField(
        max_length=20,
        required=False,
        label="Serie / Prefijo",
        widget=forms.TextInput(attrs={'placeholder': 'Ej: AU06 o WJ53', 'style': 'max-width: 250px;'}),
        help_text="Prefijo de la factura. Puedes editar las letras o números iniciales (ej: AU06, AU07, WJ53). Si se deja vacío, toma el de la base."
    )
    cantidad = forms.IntegerField(
        min_value=1,
        max_value=500,
        initial=3,
        label="Cantidad Total de Facturas",
        help_text="Número exacto de facturas a generar. El sistema respetará siempre esta cantidad."
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
    hora_inicial = forms.TimeField(
        initial="08:00",
        label="Hora Inicial (Rango Comercial)",
        widget=forms.TimeInput(attrs={'type': 'time', 'style': 'max-width: 250px;'}),
        help_text="Hora de inicio para distribuir las ventas durante el día."
    )
    hora_final = forms.TimeField(
        initial="20:00",
        label="Hora Final (Rango Comercial)",
        widget=forms.TimeInput(attrs={'type': 'time', 'style': 'max-width: 250px;'}),
        help_text="Hora de fin para distribuir las ventas durante el día."
    )
    cajero = forms.CharField(
        max_length=100,
        required=False,
        label="Nombre del Cajero / Operador",
        widget=forms.TextInput(attrs={'placeholder': 'Ej: ZAPATA MICHAELL', 'style': 'max-width: 250px;'}),
        help_text="Opcional. Nombre del operador que aparece en la tirilla."
    )
    factura_guia = forms.FileField(
        required=False,
        label="Adjuntar Factura Guía (.docx)",
        help_text="Opcional. Si adjuntas un archivo Word muestra, el sistema editará directamente sobre él respetando 100% el diseño original."
    )
    facturas_por_dia = forms.IntegerField(
        min_value=1,
        initial=3,
        label="Facturas por Día",
        help_text="Cuántas facturas se emitirán por cada día calendario (por ejemplo 3 por día)."
    )
    folio_inicial = forms.CharField(
        max_length=20,
        required=False,
        label="Folio Inicial",
        widget=forms.TextInput(attrs={'placeholder': 'Ej: 1001', 'style': 'max-width: 250px;'}),
        help_text="Si se deja vacío, se tomará el de la factura base incrementado en 1."
    )
    incremento_folio = forms.CharField(
        max_length=20,
        initial="1",
        required=False,
        label="Salto / Incremento de Folio",
        widget=forms.TextInput(attrs={'placeholder': 'Ej: 1, 7, 9 o 6-15', 'style': 'max-width: 250px;'}),
        help_text="Incremento entre cada factura. Ej: 1 (normal), 7 (de 7 en 7), 9 (de 9 en 9), 15 (de 15 en 15), o rango como 6-15."
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
def parse_salto_folio(salto_val):
    if not salto_val:
        return 1
    if isinstance(salto_val, int):
        return max(1, salto_val)
    s = str(salto_val).strip()
    if '-' in s:
        parts = s.split('-')
        try:
            low = int(parts[0].strip())
            high = int(parts[1].strip())
            import random
            return random.randint(min(low, high), max(low, high))
        except Exception:
            return 1
    try:
        return max(1, int(s))
    except Exception:
        return 1

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

def ejecutar_generacion_masiva_para_lote(lote, usuario=None, cajero_custom=None, serie_custom=None, folio_inicial=None, incremento_folio_val=1, comportamiento_vencimiento='plazo_relativo'):
    """
    Función centralizada para generar todas las facturas y productos de un lote.
    Garantiza QR dinámico, CUFE único, horas realistas distribuidas y folios consecutivos.
    """
    import random
    from django.utils import timezone
    factura_base = lote.factura_base
    if not factura_base:
        return []

    cantidad = lote.cantidad_generada or 1
    fecha_inicial = lote.fecha_inicio or timezone.localdate()
    fecha_final = lote.fecha_fin
    facturas_por_dia = lote.facturas_por_dia or 1
    hora_inicial = lote.hora_inicial or datetime.time(8, 0)
    hora_final = lote.hora_final or datetime.time(20, 0)
    serie_a_usar = serie_custom.strip().upper() if serie_custom else (lote.serie.strip().upper() if lote and getattr(lote, 'serie', None) else factura_base.serie)

    if not folio_inicial:
        folio_inicial = lote.folio_inicial if lote and getattr(lote, 'folio_inicial', None) else increment_folio(factura_base.folio, 1)

    dias_plazo = 0
    if factura_base.fecha_emision and factura_base.fecha_vencimiento:
        dias_plazo = (factura_base.fecha_vencimiento - factura_base.fecha_emision).days

    min_ini = hora_inicial.hour * 60 + hora_inicial.minute
    min_fin = hora_final.hour * 60 + hora_final.minute
    if min_fin <= min_ini:
        min_fin = min_ini + 600

    try:
        base_tr = int(re.sub(r'\D', '', factura_base.transaccion or '9981') or '9981')
    except Exception:
        base_tr = 9981

    try:
        base_rem = int(re.sub(r'\D', '', factura_base.rem or '9589') or '9589')
    except Exception:
        base_rem = 9589

    facturas_creadas = []
    current_folio = folio_inicial

    for i in range(cantidad):
        dia_offset = i // facturas_por_dia
        fecha_emision_nueva = fecha_inicial + datetime.timedelta(days=dia_offset)
        if fecha_final and fecha_emision_nueva > fecha_final:
            fecha_emision_nueva = fecha_final

        if comportamiento_vencimiento == 'plazo_relativo':
            fecha_vencimiento_nueva = fecha_emision_nueva + datetime.timedelta(days=dias_plazo) if dias_plazo > 0 else None
        elif comportamiento_vencimiento == 'misma_fecha':
            fecha_vencimiento_nueva = fecha_emision_nueva
        else:
            fecha_vencimiento_nueva = None

        # Asegurar que el folio no esté repetido
        while Factura.objects.filter(serie=serie_a_usar, folio=current_folio).exists():
            current_folio = increment_folio(current_folio, 1)

        folio_a_usar = current_folio
        salto_step = parse_salto_folio(incremento_folio_val or getattr(lote, 'salto_folio', '1') or '1')
        current_folio = increment_folio(current_folio, salto_step)

        # Cálculo de hora realista y única
        indice_en_dia = i % facturas_por_dia
        total_en_este_dia = min(facturas_por_dia, cantidad - (dia_offset * facturas_por_dia))
        if cantidad == 1:
            hora_calculada = timezone.localtime(timezone.now()).time().replace(microsecond=0)
        elif total_en_este_dia > 1:
            step = (min_fin - min_ini) / (total_en_este_dia + 1)
            base_min = min_ini + (indice_en_dia + 1) * step
            jitter = random.randint(-6, 6)
            calc_min = int(base_min + jitter)
            calc_min = max(min_ini, min(min_fin, calc_min))
            hora_calculada = datetime.time(calc_min // 60, calc_min % 60, random.randint(0, 59))
        else:
            rango = min_fin - min_ini
            offset_var = ((dia_offset * 137 + 45) % max(rango - 60, 1))
            calc_min = max(min_ini, min(min_fin, min_ini + 30 + offset_var))
            hora_calculada = datetime.time(calc_min // 60, calc_min % 60, random.randint(0, 59))

        tr_calculado = str(base_tr + i + 1)
        rem_calculado = f"REM{base_rem + i + 1:05d}"

        copia_factura = Factura(
            plantilla=lote.plantilla or getattr(factura_base, 'plantilla', 'JUMBO'),
            lote_generacion=lote,
            serie=serie_a_usar,
            folio=folio_a_usar,
            fecha_emision=fecha_emision_nueva,
            hora_emision=hora_calculada,
            fecha_vencimiento=fecha_vencimiento_nueva,
            cajero=cajero_custom or getattr(factura_base, 'cajero', 'ZAPATA MICHAELL'),
            terminal=getattr(factura_base, 'terminal', '1034'),
            registro=getattr(factura_base, 'registro', '05'),
            transaccion=tr_calculado,
            rem=rem_calculado,
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

    return facturas_creadas


@admin.register(Factura)

class FacturaAdmin(admin.ModelAdmin):
    actions = ['actualizar_codigo_qr_action', 'descargar_seleccionadas_word_unificado_action', 'descargar_seleccionadas_zip_action', 'resetear_consecutivo_ids_action']

    @admin.action(description="Sincronizar y resetear consecutivo de IDs")
    def resetear_consecutivo_ids_action(self, request, queryset):
        from .models import resetear_autoincrement_tabla
        resetear_autoincrement_tabla('facturas_factura')
        resetear_autoincrement_tabla('facturas_lotegeneracionfacturas')
        max_id = Factura.objects.order_by('-id').values_list('id', flat=True).first() or 0
        messages.success(request, f"¡Secuencia de IDs reindexada con éxito! El próximo ID a generar será #{max_id + 1}.")

    @admin.action(description="Regenerar códigos QR de seleccionadas")
    def actualizar_codigo_qr_action(self, request, queryset):
        count = 0
        for factura in queryset:
            factura.generar_qr()
            count += 1
        self.message_user(request, f"Se actualizaron exitosamente los códigos QR de {count} factura(s).", level=messages.SUCCESS)

    @admin.action(description="Descargar seleccionadas en un solo Word (.docx)")
    def descargar_seleccionadas_word_unificado_action(self, request, queryset):
        from . import word_generator
        from django.http import HttpResponse
        import copy, io
        facturas = list(queryset.order_by('id'))
        if not facturas:
            return
        doc_base = word_generator.generar_documento_factura(facturas[0])
        if len(facturas) > 1:
            try:
                from docxcompose.composer import Composer
                composer = Composer(doc_base)
                for f in facturas[1:]:
                    doc_next = word_generator.generar_documento_factura(f)
                    composer.append(doc_next)
            except Exception as e:
                print("Composer error:", e)
        stream = io.BytesIO()
        doc_base.save(stream)
        stream.seek(0)
        response = HttpResponse(
            stream.getvalue(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        response['Content-Disposition'] = f'attachment; filename="Facturas_Seleccionadas_Unificadas_{len(facturas)}.docx"'
        return response

    @admin.action(description="Descargar seleccionadas en ZIP (.zip)")
    def descargar_seleccionadas_zip_action(self, request, queryset):
        from . import word_generator
        from django.http import HttpResponse
        import io, zipfile
        zip_stream = io.BytesIO()
        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in queryset:
                f_stream = word_generator.generar_factura_word_bytes(f)
                zf.writestr(f"Factura_{f.serie}_{f.folio}.docx", f_stream.getvalue())
        zip_stream.seek(0)
        response = HttpResponse(zip_stream.getvalue(), content_type='application/zip')
        response['Content-Disposition'] = f'attachment; filename="Facturas_Seleccionadas_{queryset.count()}.zip"'
        return response


    list_display = ('id', 'tipo_origen', 'plantilla_badge', 'serie_folio', 'emisor_nombre', 'fecha_emision', 'hora_emision', 'total_formateado', 'acciones_rapidas')
    search_fields = ('id', 'cufe', 'serie', 'folio', 'emisor_nombre', 'receptor_nombre', 'emisor_nit')
    list_filter = ('plantilla', 'estado', 'fecha_emision')

    def tipo_origen(self, obj):
        if not obj.lote_generacion:
            return mark_safe('<span style="color: #94a3b8; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.5px; border: 1px solid #475569; padding: 2px 6px; border-radius: 3px;">Base</span>')
        url_lote = reverse('admin:facturas_lotegeneracionfacturas_change', args=[obj.lote_generacion.id])
        return mark_safe(f'<a href="{url_lote}" style="color: #64748b; font-size: 0.75rem; text-decoration: none;">Lote #{obj.lote_generacion.id}</a>')
    tipo_origen.short_description = 'Tipo'

    def plantilla_badge(self, obj):
        colores = {
            'JUMBO': '#22c55e',
            'EXITO': '#eab308',
        }
        text_col = colores.get(obj.plantilla, '#94a3b8')
        nombre = 'Jumbo' if obj.plantilla == 'JUMBO' else ('Éxito' if obj.plantilla == 'EXITO' else obj.plantilla)
        return mark_safe(f'<span style="color: {text_col}; font-weight: 500; font-size: 0.85rem;">{nombre}</span>')
    plantilla_badge.short_description = 'Plantilla'

    def serie_folio(self, obj):
        return f"{obj.serie}-{obj.folio}" if obj.plantilla == 'JUMBO' else f"{obj.serie}{obj.folio}"
    serie_folio.short_description = 'Nº Factura'

    def total_formateado(self, obj):
        from .word_generator import fmt_pesos
        return f"${fmt_pesos(obj.total)} COP"
    total_formateado.short_description = 'Total'

    def acciones_rapidas(self, obj):
        if obj.cufe:
            url_prev = reverse('previsualizar_tirilla', args=[obj.cufe])
            url_word = reverse('descargar_word_factura', args=[obj.cufe])
            return mark_safe(
                f'<a href="{url_prev}" target="_blank" style="color: #94a3b8; border: 1px solid #475569; padding: 2px 7px; border-radius: 3px; font-size: 0.75rem; text-decoration: none; margin-right: 5px;">Vista Previa</a>'
                f'<a href="{url_word}" style="color: #94a3b8; border: 1px solid #475569; padding: 2px 7px; border-radius: 3px; font-size: 0.75rem; text-decoration: none;">Word</a>'
            )
        return "-"
    acciones_rapidas.short_description = 'Acciones'
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
        ('Datos de Caja y Facturación POS (Tirilla)', {'fields': ('hora_emision', 'cajero', 'terminal', 'registro', 'transaccion', 'rem')}),
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
                '<a class="button" href="{}" style="background-color: #334155; color: #f1f5f9; border: 1px solid #475569; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-size: 0.8rem; font-weight: 500;">Descargar Word (80mm)</a>',
                url
            )
        return "Guarde la factura primero para generar el Word"
    descargar_word.short_description = 'Descargar Word 80mm'

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                'regenerar-todos-los-qr/',
                self.admin_site.admin_view(self.regenerar_todos_qr_view),
                name='facturas_factura_regenerar_todos_qr',
            ),
            path(
                '<path:object_id>/generar-masivo/',
                self.admin_site.admin_view(self.generar_masivo_view),
                name='facturas_factura_generar_masivo',
            ),
        ]
        return custom_urls + urls

    def regenerar_todos_qr_view(self, request):
        count = 0
        for factura in Factura.objects.all():
            factura.generar_qr()
            count += 1
        self.message_user(request, f"¡Éxito! Se regeneraron los códigos QR de todas las {count} facturas con el nuevo dominio.", level=messages.SUCCESS)
        return redirect('admin:facturas_factura_changelist')

    # Vista de generación masiva adaptada para auditar y registrar el lote
    def generar_masivo_view(self, request, object_id):
        factura_base = get_object_or_404(Factura, pk=object_id)
        error_message = None

        if request.method == 'POST':
            form = GenerarFacturasMasivasForm(request.POST, request.FILES)
            if form.is_valid():
                cantidad = form.cleaned_data['cantidad']
                fecha_inicial = form.cleaned_data['fecha_inicial']
                fecha_final = form.cleaned_data['fecha_final']
                facturas_por_dia = form.cleaned_data['facturas_por_dia']
                folio_inicial = form.cleaned_data['folio_inicial']
                incremento_folio_val = form.cleaned_data['incremento_folio']
                comportamiento_vencimiento = form.cleaned_data['comportamiento_vencimiento']
                hora_inicial = form.cleaned_data.get('hora_inicial') or datetime.time(8, 30)
                hora_final = form.cleaned_data.get('hora_final') or datetime.time(20, 0)
                cajero_custom = form.cleaned_data.get('cajero')
                factura_guia_file = request.FILES.get('factura_guia')

                plantilla_sel = form.cleaned_data.get('plantilla') or factura_base.plantilla
                base_a_usar = factura_base
                if plantilla_sel != factura_base.plantilla:
                    otra_base = Factura.objects.filter(plantilla=plantilla_sel, lote_generacion__isnull=True).first() or Factura.objects.filter(plantilla=plantilla_sel).first()
                    if otra_base:
                        base_a_usar = otra_base
                        if not folio_inicial:
                            folio_inicial = increment_folio(base_a_usar.folio, 1)

                # Respetar siempre la cantidad exacta solicitada por el usuario
                if not cantidad:
                    if fecha_final and fecha_final >= fecha_inicial:
                        dias = (fecha_final - fecha_inicial).days + 1
                        cantidad = dias * (facturas_por_dia or 1)
                    else:
                        cantidad = 3

                try:
                    with transaction.atomic():
                        fecha_fin_calculada = fecha_final or (fecha_inicial + datetime.timedelta(days=((cantidad - 1) // max(facturas_por_dia, 1))))
                        lote = LoteGeneracionFacturas.objects.create(
                            plantilla=plantilla_sel,
                            usuario_creador=request.user,
                            factura_base=base_a_usar,
                            cantidad_generada=cantidad,
                            fecha_inicio=fecha_inicial,
                            fecha_fin=fecha_fin_calculada,
                            facturas_por_dia=facturas_por_dia,
                            hora_inicial=hora_inicial,
                            hora_final=hora_final,
                            factura_guia=factura_guia_file,
                            observaciones=f"Generación masiva ({plantilla_sel}) de {cantidad} facturas.",
                            estado='ACTIVO'
                        )
                        serie_custom = form.cleaned_data.get('serie')
                        facturas_creadas = ejecutar_generacion_masiva_para_lote(
                            lote=lote,
                            usuario=request.user,
                            cajero_custom=cajero_custom,
                            serie_custom=serie_custom,
                            folio_inicial=folio_inicial,
                            incremento_folio_val=incremento_folio_val,
                            comportamiento_vencimiento=comportamiento_vencimiento
                        )

                        folios_creados = [f"{f.serie}{f.folio}" for f in facturas_creadas]
                        rango_folios = f"desde {folios_creados[0]} hasta {folios_creados[-1]}" if len(folios_creados) > 1 else folios_creados[0]
                        messages.success(
                            request,
                            f"Se generaron con éxito {len(facturas_creadas)} facturas masivas ({rango_folios}) en el Lote #{lote.id}."
                        )
                        return redirect('admin:facturas_lotegeneracionfacturas_change', lote.pk)

                except forms.ValidationError as e:
                    error_message = e.message
                except Exception as e:
                    error_message = f"Error crítico durante el procesamiento en base de datos: {str(e)}"

        else:
            siguiente_folio = increment_folio(factura_base.folio, 1)
            form = GenerarFacturasMasivasForm(initial={
                'plantilla': factura_base.plantilla,
                'serie': factura_base.serie,
                'cantidad': 3,
                'facturas_por_dia': 3,
                'fecha_inicial': timezone.localdate(),
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
    list_display = ('id', 'plantilla_badge', 'usuario_creador', 'fecha_creacion', 'factura_base_link', 'cantidad_generada', 'estado_badge', 'acciones_descarga')

    def plantilla_badge(self, obj):
        colores = {
            'JUMBO': ('#dcfce7', '#166534'),
            'EXITO': ('#fef3c7', '#92400e'),
        }
        bg, text_col = colores.get(obj.plantilla, ('#f1f5f9', '#475569'))
        return mark_safe(f'<span style="background: {bg}; color: {text_col}; font-weight: bold; padding: 2px 8px; border-radius: 4px;">{obj.get_plantilla_display().split(" ")[0]}</span>')
    plantilla_badge.short_description = 'Plantilla'

    def plantilla_badge(self, obj):
        nombre = 'Jumbo' if obj.plantilla == 'JUMBO' else ('Éxito' if obj.plantilla == 'EXITO' else obj.plantilla)
        colores = {'JUMBO': '#22c55e', 'EXITO': '#eab308'}
        return mark_safe(f'<span style="color: {colores.get(obj.plantilla, "#94a3b8")}; font-weight: 500;">{nombre}</span>')
    plantilla_badge.short_description = 'Plantilla'

    def acciones_descarga(self, obj):
        if obj.estado == 'ACTIVO' and obj.cantidad_generada > 0:
            url_word = reverse('descargar_lote_word_unificado', args=[obj.pk])
            url_zip = reverse('descargar_lote_zip', args=[obj.pk])
            return mark_safe(
                f'<a href="{url_word}" style="color: #94a3b8; border: 1px solid #475569; padding: 2px 7px; border-radius: 3px; font-size: 0.75rem; text-decoration: none; margin-right: 5px;">Word</a>'
                f'<a href="{url_zip}" style="color: #94a3b8; border: 1px solid #475569; padding: 2px 7px; border-radius: 3px; font-size: 0.75rem; text-decoration: none;">ZIP</a>'
            )
        return "-"
    acciones_descarga.short_description = 'Descargar'
    list_filter = ('estado', 'fecha_creacion')
    search_fields = ('id', 'usuario_creador__username', 'factura_base__serie', 'factura_base__folio')
    change_form_template = 'admin/facturas/lotegeneracionfacturas/change_form.html'

    def get_readonly_fields(self, request, obj=None):
        if obj is None:
            return ()
        return (
            'plantilla', 'usuario_creador', 'usuario_cancelador', 'fecha_creacion', 'fecha_cancelacion', 
            'factura_base', 'serie', 'folio_inicial', 'salto_folio', 'cantidad_generada', 'fecha_inicio', 'fecha_fin', 
            'facturas_por_dia', 'hora_inicial', 'hora_final', 'factura_guia', 'observaciones', 'motivo_cancelacion', 'estado', 
            'mostrar_estadisticas', 'listado_facturas'
        )

    def get_fieldsets(self, request, obj=None):
        if obj is None:
            return (
                ('Configuración de Generación Masiva', {
                    'description': 'Selecciona la plantilla, fechas y horas. Si dejas la fecha de fin vacía, se generará la cantidad exacta indicada. Puedes cambiar el prefijo (Serie) antes de producir en serie.',
                    'fields': (
                        'plantilla',
                        'factura_base',
                        ('serie', 'folio_inicial', 'salto_folio'),
                        'cantidad_generada',
                        ('fecha_inicio', 'fecha_fin'),
                        'facturas_por_dia',
                        ('hora_inicial', 'hora_final'),
                        'factura_guia',
                        'observaciones',
                    )
                }),
            )
        return (
            ('Información de Registro', {
                'fields': ('plantilla', 'usuario_creador', 'fecha_creacion', 'factura_base', 'estado')
            }),
            ('Detalle de Generación Masiva', {
                'fields': (('serie', 'folio_inicial', 'salto_folio'), 'cantidad_generada', 'fecha_inicio', 'fecha_fin', 'facturas_por_dia', 'hora_inicial', 'hora_final', 'factura_guia', 'observaciones')
            }),
            ('Resumen del Lote', {
                'fields': ('mostrar_estadisticas',)
            }),
            ('Documentos del Lote', {
                'fields': ('listado_facturas',)
            }),
            ('Auditoría', {
                'fields': ('usuario_cancelador', 'fecha_cancelacion', 'motivo_cancelacion')
            }),
        )

    def formfield_for_foreignkey(self, db_field, request, **kwargs):
        if db_field.name == "factura_base":
            kwargs["initial"] = Factura.objects.filter(lote_generacion__isnull=True).order_by('id').first()
        return super().formfield_for_foreignkey(db_field, request, **kwargs)

    def save_model(self, request, obj, form, change):
        if not change:
            if not obj.usuario_creador:
                obj.usuario_creador = request.user
            obj.estado = 'ACTIVO'

            # Asegurar plantilla
            if not obj.plantilla:
                obj.plantilla = 'JUMBO'

            # Enlazar o corregir factura base acorde a la plantilla
            if not obj.factura_base:
                obj.factura_base = Factura.objects.filter(plantilla=obj.plantilla, lote_generacion__isnull=True).first() or Factura.objects.filter(plantilla=obj.plantilla).first() or Factura.objects.order_by('-id').first()
            elif obj.factura_base.plantilla != obj.plantilla:
                matching_base = Factura.objects.filter(plantilla=obj.plantilla, lote_generacion__isnull=True).first() or Factura.objects.filter(plantilla=obj.plantilla).first()
                if matching_base:
                    obj.factura_base = matching_base

            if not obj.fecha_inicio:
                obj.fecha_inicio = timezone.localdate()

            if not obj.facturas_por_dia:
                obj.facturas_por_dia = 3

            # Respetar SIEMPRE la cantidad_generada exacta solicitada por el usuario
            if not obj.cantidad_generada:
                if obj.fecha_inicio and obj.fecha_fin and obj.fecha_fin >= obj.fecha_inicio:
                    dias = (obj.fecha_fin - obj.fecha_inicio).days + 1
                    obj.cantidad_generada = dias * (obj.facturas_por_dia or 1)
                else:
                    obj.cantidad_generada = 3

            if not obj.fecha_fin:
                dias = (obj.cantidad_generada - 1) // max(obj.facturas_por_dia or 1, 1)
                obj.fecha_fin = obj.fecha_inicio + datetime.timedelta(days=dias)

            super().save_model(request, obj, form, change)

            if obj.factura_base and obj.cantidad_generada > 0:
                with transaction.atomic():
                    facturas_creadas = ejecutar_generacion_masiva_para_lote(
                        lote=obj, 
                        usuario=request.user,
                        serie_custom=obj.serie,
                        folio_inicial=obj.folio_inicial,
                        incremento_folio_val=obj.salto_folio
                    )
                    messages.success(
                        request,
                        f"Lote #{obj.id} ({obj.plantilla}) generado con éxito. Se crearon {len(facturas_creadas)} facturas."
                    )
        else:
            super().save_model(request, obj, form, change)

    def response_add(self, request, obj, post_url_continue=None):
        if '_continue' in request.POST:
            return redirect('admin:facturas_lotegeneracionfacturas_change', obj.pk)
        elif '_addanother' in request.POST:
            return redirect('admin:facturas_lotegeneracionfacturas_add')
        return redirect('admin:facturas_lotegeneracionfacturas_change', obj.pk)


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
        
        from .word_generator import fmt_pesos
        html = f"""
        <div style="display: flex; flex-wrap: wrap; gap: 24px; padding: 10px 0; font-size: 0.85rem; color: #94a3b8;">
            <div><span style="color: #64748b; font-size: 0.75rem; text-transform: uppercase;">Total copias:</span> <strong style="color: inherit; margin-left: 4px;">{total_facturas}</strong></div>
            <div><span style="color: #64748b; font-size: 0.75rem; text-transform: uppercase;">Monto total:</span> <strong style="color: inherit; margin-left: 4px;">${fmt_pesos(suma_facturada)} COP</strong></div>
            <div><span style="color: #64748b; font-size: 0.75rem; text-transform: uppercase;">Rango:</span> <strong style="color: inherit; margin-left: 4px;">{fecha_ini.strftime('%d/%m/%Y')} &mdash; {fecha_fin.strftime('%d/%m/%Y') if fecha_fin else '-'}</strong></div>
            <div><span style="color: #64748b; font-size: 0.75rem; text-transform: uppercase;">Frecuencia:</span> <strong style="color: inherit; margin-left: 4px;">{obj.facturas_por_dia} por día</strong></div>
        </div>
        """
        return mark_safe(html)
    mostrar_estadisticas.short_description = 'Resumen del Lote'

    def listado_facturas(self, obj):
        if not obj or not obj.pk:
            return "-"
        facturas = obj.facturas.all()
        if not facturas.exists():
            return "Sin facturas."
        from .word_generator import fmt_pesos
        
        html = """
        <div style="max-height: 280px; overflow-y: auto; border: 1px solid #334155; border-radius: 4px;">
        <table style="width:100%; border-collapse:collapse; font-size:0.8rem; text-align:left;">
            <thead>
                <tr style="border-bottom:1px solid #334155; color:#94a3b8;">
                    <th style="padding:6px 10px;">Nº Factura</th>
                    <th style="padding:6px 10px;">Fecha</th>
                    <th style="padding:6px 10px;">Hora</th>
                    <th style="padding:6px 10px;">Total</th>
                    <th style="padding:6px 10px; text-align:center;">Acción</th>
                </tr>
            </thead>
            <tbody>
        """
        for f in facturas:
            url = reverse('admin:facturas_factura_change', args=[f.pk])
            num = f"{f.serie}-{f.folio}" if f.plantilla == 'JUMBO' else f"{f.serie}{f.folio}"
            html += f"""
                <tr style="border-bottom:1px solid rgba(255,255,255,0.05); color:#cbd5e1;">
                    <td style="padding:6px 10px; font-weight:500;">{num}</td>
                    <td style="padding:6px 10px;">{f.fecha_emision.strftime('%d/%m/%Y') if f.fecha_emision else '-'}</td>
                    <td style="padding:6px 10px;">{f.hora_emision.strftime('%H:%M') if f.hora_emision else '-'}</td>
                    <td style="padding:6px 10px;">${fmt_pesos(f.total)}</td>
                    <td style="padding:6px 10px; text-align:center;">
                        <a href="{url}" target="_blank" style="color:#94a3b8; border:1px solid #475569; padding:2px 7px; border-radius:3px; font-size:0.75rem; text-decoration:none;">Ver</a>
                    </td>
                </tr>
            """
        html += "</tbody></table></div>"
        return mark_safe(html)
    listado_facturas.short_description = 'Documentos del Lote'

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
            motivo = request.POST.get('motivo', '').strip() or 'Cancelación por usuario'
            if total_con_riesgo > 0 and not request.user.is_superuser:
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
            motivo = request.POST.get('motivo', '').strip() or 'Desactivación por usuario'
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
                        f"Lote #{lote.id} desactivado en modo seguro. Las {total_anuladas} facturas se conservan intactas en la base de datos pero marcadas como 'Anuladas'."
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
