from django.contrib import admin
from .models import Rule

@admin.register(Rule)
class RuleAdmin(admin.ModelAdmin):
    list_display = (
        'name',
        'account',
        'imap_folder',
        'enabled',
        'from_contains',
        'subject_contains',
        'dest_folder',
        'created_at',
    )
    list_filter = ('enabled', 'account', 'imap_folder')
    search_fields = ('name', 'from_contains', 'subject_contains', 'dest_folder')
