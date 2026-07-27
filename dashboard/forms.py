from django import forms
from django.contrib.auth.models import User
from django.contrib.auth.forms import AuthenticationForm
from .models import Barang, Kustomer, Supplier, APISetting, CompanyProfile
from django import forms
from .models import UserProfile


class LoginForm(AuthenticationForm):
    username = forms.CharField(widget=forms.TextInput(attrs={'class': 'form-control'}))
    password = forms.CharField(widget=forms.PasswordInput(attrs={'class': 'form-control'}))

class BarangForm(forms.ModelForm):
    class Meta:
        model = Barang
        fields = ['id_barang', 'nama_barang', 'lot', 'kategori', 'deskripsi']
        widgets = {
            'id_barang': forms.TextInput(attrs={'class': 'form-control'}),
            'nama_barang': forms.TextInput(attrs={'class': 'form-control'}),
            'lot': forms.TextInput(attrs={'class': 'form-control'}),
            'kategori': forms.TextInput(attrs={'class': 'form-control'}),
            'deskripsi': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
        }

class KustomerForm(forms.ModelForm):
    class Meta:
        model = Kustomer
        fields = ['id_kustomer', 'nama_kustomer', 'alamat', 'telepon', 'email']
        widgets = {
            'id_kustomer': forms.TextInput(attrs={'class': 'form-control'}),
            'nama_kustomer': forms.TextInput(attrs={'class': 'form-control'}),
            'alamat': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'telepon': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
        }

class SupplierForm(forms.ModelForm):
    class Meta:
        model = Supplier
        fields = ['id_supplier', 'nama_supplier', 'alamat', 'telepon', 'email']
        widgets = {
            'id_supplier': forms.TextInput(attrs={'class': 'form-control'}),
            'nama_supplier': forms.TextInput(attrs={'class': 'form-control'}),
            'alamat': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'telepon': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
        }

class APISettingForm(forms.ModelForm):
    class Meta:
        model = APISetting
        fields = ['write_api_key', 'read_api_key', 'channel_id', 'api_url']
        widgets = {
            'write_api_key': forms.TextInput(attrs={'class': 'form-control'}),
            'read_api_key': forms.TextInput(attrs={'class': 'form-control'}),
            'channel_id': forms.TextInput(attrs={'class': 'form-control'}),
            'api_url': forms.URLInput(attrs={'class': 'form-control'}),
        }

class CompanyProfileForm(forms.ModelForm):
    class Meta:
        model = CompanyProfile
        fields = ['name', 'address', 'phone', 'email', 'logo']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'address': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'phone': forms.TextInput(attrs={'class': 'form-control'}),
            'email': forms.EmailInput(attrs={'class': 'form-control'}),
            'logo': forms.FileInput(attrs={'class': 'form-control'}),
            'website': forms.URLInput(attrs={'class': 'forms-control', 'placeholder': 'https://www.example.com'}),
            
        }
 # dashboard/forms.py


class UserProfileForm(forms.ModelForm):
    class Meta:
        model = UserProfile
        fields = ['avatar', 'phone', 'address', 'bio', 'birth_date', 'gender']
        widgets = {
            'bio': forms.Textarea(attrs={'rows': 3}),
            'address': forms.Textarea(attrs={'rows': 3}),
            'birth_date': forms.DateInput(attrs={'type': 'date'}),
        }      