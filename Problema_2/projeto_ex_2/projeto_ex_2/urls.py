
from django.urls import path
from app_projeto import views
from django.contrib import admin

urlpatterns = [
    path('', views.home, name='home'),
    path('usuarios/', views.cadastro, name='listagem_contaminantes'),
    path('admin/', admin.site.urls),
]
