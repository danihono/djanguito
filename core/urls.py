
from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('analise/', views.analise_mercado, name='analise_mercado'),
]
