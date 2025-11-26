from django.contrib import admin
from .models import Rule, ProcessedEmail, JobExecution, AttachmentLog

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

@admin.register(ProcessedEmail)
class ProcessedEmailAdmin(admin.ModelAdmin):
    list_display = ('account', 'rule', 'folder', 'uid', 'processed_at')
    list_filter = ('account', 'rule', 'folder')
    search_fields = ('uid', 'message_id_header')


@admin.register(JobExecution)
class JobExecutionAdmin(admin.ModelAdmin):
    list_display = (
        'rule',
        'status',
        'emails_processed',
        'attachments_downloaded',
        'started_at',
        'finished_at',
    )
    list_filter = ('status', 'rule')
    search_fields = ('rule__name', 'error_message')


@admin.register(AttachmentLog)
class AttachmentLogAdmin(admin.ModelAdmin):
    list_display = (
        'file_name',
        'rule',
        'account',
        'folder',
        'size_bytes',
        'downloaded_at',
    )
    list_filter = ('rule', 'account', 'folder')
    search_fields = ('file_name', 'original_filename', 'sender', 'subject', 'message_id')