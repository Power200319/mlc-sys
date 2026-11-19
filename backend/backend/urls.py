"""
URL configuration for backend project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/6.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path

from app import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('api/health/', views.health, name='health'),
    path('api/grades/', views.grades, name='grades'),
    path('api/early-years/profiles/', views.early_years_profiles_view, name='early-years-profiles'),
    path('api/early-years/process/', views.process_early_years, name='early-years-process'),
    path('api/early-years/download/<str:job_id>/', views.download_early_years, name='early-years-download'),
    path('api/process/', views.process, name='process'),
    path('api/generate-reports/', views.generate_reports, name='generate-reports'),
    path('api/download/<str:job_id>/<str:kind>/', views.download, name='download'),
]
