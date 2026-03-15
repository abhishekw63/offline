from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    path('', views.HomeView.as_view(), name='home'),
    path('login/', views.HomeView.as_view(), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='/'), name='logout'),
    path('departments/', views.DepartmentsView.as_view(), name='departments'),
    path('profile/', views.ProfileView.as_view(), name='profile'),
    path('password-change/', views.CustomPasswordChangeView.as_view(), name='password_change'),
]