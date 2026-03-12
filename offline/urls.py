from django.urls import path, include
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    path('', views.HomeView.as_view(), name='home'),
    path('login/', views.HomeView.as_view(), name='login'), # Alias for form target
    path('logout/', auth_views.LogoutView.as_view(next_page='/'), name='logout'),
    path('dashboard/', views.IndexView.as_view(), name='index'),
    path('process/', views.ProcessFilesView.as_view(), name='process_files'),
]
