from django.shortcuts import render
from django.views.generic import TemplateView
from django.contrib.auth.views import LoginView
from django.contrib.auth.mixins import LoginRequiredMixin
from django.urls import reverse_lazy

class HomeView(LoginView):
    template_name = 'core/home.html'
    redirect_authenticated_user = False

    def get_success_url(self):
        return reverse_lazy('departments')

class DepartmentsView(LoginRequiredMixin, TemplateView):
    template_name = 'core/departments.html'
