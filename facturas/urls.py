from django.urls import path
from . import views

urlpatterns = [
    path('', views.buscar_factura, name='buscar_factura'),
    path('factura/<str:cufe>/', views.ver_factura, name='ver_factura'),
    path('factura/<str:cufe>/pdf/', views.descargar_pdf_factura, name='descargar_pdf_factura'),
    path('factura/<str:cufe>/word/', views.descargar_word_factura, name='descargar_word_factura'),
    path('factura/<str:cufe>/preview/', views.previsualizar_tirilla, name='previsualizar_tirilla'),
    path('lote/<int:lote_id>/word-unificado/', views.descargar_lote_word_unificado, name='descargar_lote_word_unificado'),
    path('lote/<int:lote_id>/zip/', views.descargar_lote_zip, name='descargar_lote_zip'),
    path('crear-admin/', views.crear_admin),
    path('api/emisor-data/', views.buscar_datos_emisor, name='buscar_datos_emisor'),
]
