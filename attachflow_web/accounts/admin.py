from django.contrib import admin, messages
from django.http import HttpResponseRedirect

from .models import EmailAccount
from attachflow_core.email_core import (
    EmailAccountConfig,
    test_imap_connection,
)


@admin.register(EmailAccount)
class EmailAccountAdmin(admin.ModelAdmin):
    """
    Admin de Contas de Email com botão 'Testar ligação IMAP'
    na página de edição da conta.
    """

    # Template customizado com o botão extra
    change_form_template = "admin/accounts/emailaccount/change_form.html"

    list_display = (
        "name",
        "username",
        "host",
        "folder",
        "active",
        "use_ssl",
        "last_connection_ok",
        "cached_folders_count",
        "created_at",
    )
    list_filter = ("active", "use_ssl", "host", "last_connection_ok")
    search_fields = ("name", "username", "host", "folder")

    readonly_fields = (
        "cached_folders",
        "last_connection_ok",
        "last_connection_message",
        "created_at",
        "updated_at",
    )

    def cached_folders_count(self, obj):
        """
        Número de pastas IMAP descobertas no último teste.
        """
        if not obj.cached_folders:
            return 0
        return len(obj.cached_folders)

    cached_folders_count.short_description = "Pastas IMAP"

    def response_change(self, request, obj, **kwargs):
        """
        Intercepta o submit do botão 'Testar ligação IMAP'.

        - O Django primeiro grava as alterações da conta
        - Depois entra aqui
        - Se o POST tiver o botão especial, executamos o teste
        - Actualizamos os campos de estado
        - E voltamos para a mesma página
        """
        if "_test_connection" in request.POST:
            cfg = EmailAccountConfig(
                host=obj.host,
                port=obj.port,
                username=obj.username,
                password=obj.password,
                use_ssl=obj.use_ssl,
                folder=obj.folder,
                name=obj.name,
            )

            try:
                result = test_imap_connection(cfg)
            except Exception as e:
                messages.error(
                    request,
                    f"[{obj.name}] Erro inesperado ao testar ligação: {e}",
                )
                return HttpResponseRedirect(".")

            # Actualizar estado na BD
            obj.cached_folders = result.folders
            obj.last_connection_ok = result.success
            obj.last_connection_message = result.message
            obj.save(
                update_fields=[
                    "cached_folders",
                    "last_connection_ok",
                    "last_connection_message",
                ]
            )

            # Mensagens amigáveis
            if result.success:
                messages.success(
                    request,
                    f"[{obj.name}] Ligação IMAP OK. "
                    f"{len(result.folders)} pastas encontradas.",
                )
            else:
                messages.error(
                    request,
                    f"[{obj.name}] Falha na ligação IMAP: {result.message}",
                )

            return HttpResponseRedirect(".")

        # Outros botões (Guardar, Guardar e continuar, etc.)
        return super().response_change(request, obj, **kwargs)
