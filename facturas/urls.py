from django.urls import path
from . import views

urlpatterns = [
    path('', views.buscar_factura, name='buscar_factura'),
    path('factura/<str:cufe>/', views.ver_factura, name='ver_factura'),
]
