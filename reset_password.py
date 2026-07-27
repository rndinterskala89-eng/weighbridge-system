# reset_password.py
import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'weight_monitoring.settings')
django.setup()

from django.contrib.auth.models import User

def reset_password(username, new_password):
    try:
        user = User.objects.get(username=username)
        user.set_password(new_password)
        user.save()
        print(f"✅ Password untuk user '{username}' berhasil direset!")
        return True
    except User.DoesNotExist:
        print(f"❌ User '{username}' tidak ditemukan!")
        return False

if __name__ == "__main__":
    username = input("Masukkan username: ")
    new_password = input("Masukkan password baru: ")
    reset_password(username, new_password)