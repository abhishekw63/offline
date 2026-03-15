from django.shortcuts import render
from django.views.generic import TemplateView, UpdateView
from django.contrib.auth.views import LoginView, PasswordChangeView
from django.contrib.auth.mixins import LoginRequiredMixin
from django.urls import reverse_lazy
from django.contrib import messages
from django.contrib.auth.models import User

class HomeView(LoginView):
    template_name = 'core/home.html'
    redirect_authenticated_user = False

    def get_success_url(self):
        return reverse_lazy('departments')

    def form_valid(self, form):
        messages.success(self.request, 'Login successful')
        return super().form_valid(form)

class DepartmentsView(LoginRequiredMixin, TemplateView):
    template_name = 'core/departments.html'

class ProfileView(LoginRequiredMixin, UpdateView):
    model = User
    fields = ['first_name', 'last_name', 'email']
    template_name = 'core/profile.html'
    success_url = reverse_lazy('profile')

    def get_object(self):
        return self.request.user

    def form_valid(self, form):
        messages.success(self.request, 'Your profile details were updated successfully.')
        return super().form_valid(form)

class CustomPasswordChangeView(LoginRequiredMixin, PasswordChangeView):
    template_name = 'core/password_change.html'
    success_url = reverse_lazy('profile')

    def form_valid(self, form):
        messages.success(self.request, 'Your password was successfully updated.')
        return super().form_valid(form)
