from django.contrib import admin

from .models import EmailAccount

@admin.register(EmailAccount)
class EmailAccountAdmin(admin.ModelAdmin):
    list_display = ('name', 'username', 'host', 'folder', 'active', 'use_ssl', 'created_at')
    list_filter = ('active', 'use_ssl', 'host')
    search_fields = ('name', 'username', 'host', 'folder')
