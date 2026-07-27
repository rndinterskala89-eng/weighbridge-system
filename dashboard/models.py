# dashboard/models.py
from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
from django.db.models.signals import post_save
from django.dispatch import receiver

class CompanyProfile(models.Model):
    name = models.CharField(max_length=200)
    address = models.TextField()
    phone = models.CharField(max_length=20)
    email = models.EmailField()
    logo = models.ImageField(upload_to='company_logos/', null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    website = models.URLField(blank=True, null=True, max_length=200) 

    def __str__(self):
        return self.name or "Company Profiles"
    

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        verbose_name = 'Company Profile'
        verbose_name_plural = 'Company Profiles'


class WeightData(models.Model):
    entry_id = models.IntegerField(unique=True)
    weight = models.DecimalField(max_digits=10, decimal_places=5)
    latitude = models.DecimalField(max_digits=10, decimal_places=6)
    longitude = models.DecimalField(max_digits=10, decimal_places=6)
    created_at = models.DateTimeField()
    captured_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"Weight {self.weight} at {self.created_at}"

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        ordering = ['-created_at']
        verbose_name = 'Weight Data'
        verbose_name_plural = 'Weight Data'


class Barang(models.Model):
    id_barang = models.CharField(max_length=50, unique=True, primary_key=True)
    nama_barang = models.CharField(max_length=200)
    lot = models.CharField(max_length=50)
    kategori = models.CharField(max_length=100, blank=True)
    deskripsi = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.nama_barang} - {self.id_barang}"

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        ordering = ['nama_barang']
        verbose_name = 'Barang'
        verbose_name_plural = 'Barang'


class Kustomer(models.Model):
    id_kustomer = models.CharField(max_length=50, unique=True, primary_key=True)
    nama_kustomer = models.CharField(max_length=200)
    alamat = models.TextField()
    telepon = models.CharField(max_length=20)
    email = models.EmailField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.nama_kustomer

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        ordering = ['nama_kustomer']
        verbose_name = 'Kustomer'
        verbose_name_plural = 'Kustomer'


class Supplier(models.Model):
    id_supplier = models.CharField(max_length=50, unique=True, primary_key=True)
    nama_supplier = models.CharField(max_length=200)
    alamat = models.TextField()
    telepon = models.CharField(max_length=20)
    email = models.EmailField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.nama_supplier

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        ordering = ['nama_supplier']
        verbose_name = 'Supplier'
        verbose_name_plural = 'Suppliers'


class APISetting(models.Model):
    write_api_key = models.CharField(max_length=100)
    read_api_key = models.CharField(max_length=100)
    channel_id = models.CharField(max_length=50)
    api_url = models.URLField(default='https://api.thingspeak.com/channels')
    is_active = models.BooleanField(default=True)
    last_test = models.DateTimeField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"API Setting - {self.channel_id}"

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        verbose_name = 'API Setting'
        verbose_name_plural = 'API Settings'


class UserActivity(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    action = models.CharField(max_length=200)
    details = models.TextField(blank=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.user.username} - {self.action} at {self.timestamp}"

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        ordering = ['-timestamp']
        verbose_name = 'User Activity'
        verbose_name_plural = 'User Activities'


class UserProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='userprofile')
    avatar = models.ImageField(upload_to='avatars/', null=True, blank=True, 
                               help_text="Upload foto profil user")
    phone = models.CharField(max_length=20, blank=True, help_text="Nomor telepon")
    address = models.TextField(blank=True, help_text="Alamat user")
    bio = models.TextField(blank=True, help_text="Bio singkat")
    birth_date = models.DateField(null=True, blank=True)
    gender = models.CharField(
        max_length=10, 
        choices=[('M', 'Male'), ('F', 'Female'), ('O', 'Other')],
        blank=True,
        default='M'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return f"Profile of {self.user.username}"
    
    @property
    def get_avatar_url(self):
        if self.avatar:
            return self.avatar.url
        name = self.user.get_full_name() or self.user.username
        return f"https://ui-avatars.com/api/?name={name}&background=F57C00&color=ffffff&size=72&bold=true"

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        verbose_name = 'User Profile'
        verbose_name_plural = 'User Profiles'


class Report(models.Model):
    report_type = models.CharField(max_length=50, choices=[
        ('Transaction', 'Transaction Report'),
        ('Weight', 'Weight Report'),
        ('Custom', 'Custom Report')
    ], default='Transaction')
    generated_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='reports')
    start_date = models.DateTimeField()
    end_date = models.DateTimeField()
    data = models.JSONField(default=dict)
    file_path = models.TextField(blank=True, help_text="Path ke file snapshot atau base64")
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.report_type} - {self.created_at.strftime('%Y-%m-%d')}"

    class Meta:
        app_label = 'dashboard'  # <-- TAMBAHKAN INI
        ordering = ['-created_at']
        verbose_name = 'Report'
        verbose_name_plural = 'Reports'


# ============================================
# SIGNAL - Buat UserProfile otomatis
# ============================================

@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    if created:
        UserProfile.objects.create(user=instance)
    else:
        if not hasattr(instance, 'userprofile'):
            UserProfile.objects.create(user=instance)

@receiver(post_save, sender=User)
def save_user_profile(sender, instance, **kwargs):
    if hasattr(instance, 'userprofile'):
        instance.userprofile.save()