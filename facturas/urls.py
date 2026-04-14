from django.urls import path
from . import views

urlpatterns = [
    path('', views.buscar_factura, name='buscar_factura'),
    path('factura/<str:cufe>/', views.ver_factura, name='ver_factura'),
    path('factura/<str:cufe>/pdf/', views.descargar_pdf_factura, name='descargar_pdf_factura'),
    path('factura/<str:cufe>/word/', views.descargar_word_factura, name='descargar_word_factura'),
    path('crear-admin/', views.crear_admin),
]

