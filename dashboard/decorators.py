# dashboard/decorators.py
from django.contrib.auth import logout
from django.shortcuts import redirect
from django.http import JsonResponse
from functools import wraps
from .models import UserActivity

def password_required(view_func):
    @wraps(view_func)
    def wrapper(request, *args, **kwargs):
        if request.method == 'POST':
            password = request.POST.get('password')
            if not password or not request.user.check_password(password):
                return JsonResponse({
                    'status': 'error',
                    'message': 'Password required for this action'
                })
        return view_func(request, *args, **kwargs)
    return wrapper

def log_activity(view_func):
    @wraps(view_func)
    def wrapper(request, *args, **kwargs):
        response = view_func(request, *args, **kwargs)
        
        # Log GET requests to pages
        if request.method == 'GET' and request.user.is_authenticated:
            UserActivity.objects.create(
                user=request.user,
                action=f'View: {request.path}',
                ip_address=request.META.get('REMOTE_ADDR')
            )
        elif request.method == 'POST' and request.user.is_authenticated:
            UserActivity.objects.create(
                user=request.user,
                action=f'Action: {request.path}',
                ip_address=request.META.get('REMOTE_ADDR')
            )
        
        return response
    return wrapper

def check_user_active(view_func):
    @wraps(view_func)
    def wrapper(request, *args, **kwargs):
        if request.user.is_authenticated and not request.user.is_active:
            logout(request)
            return redirect('login')
        return view_func(request, *args, **kwargs)
    return wrapper