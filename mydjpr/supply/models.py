from django.contrib.auth.models import User
from django.db import models

class UserProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    file_path_user = models.CharField(max_length=255)

    def __str__(self):
        return f'{self.user.username} - {self.file_path_user}'