from django.contrib import admin
from django.urls import path
import awsApp.views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', awsApp.views.index, name='index'),
]
