from django.db import models

class EmailAccount(models.Model):
    """
    Representa uma conta de email que o AttachFlow vai monitorizar via IMAP.
    """

    name = models.CharField(
        max_length=100,
        help_text="Nome amigável da conta (ex.: Outlook Pessoal, Conta LBC, etc.)"
    )
    host = models.CharField(
        max_length=255,
        default="outlook.office365.com",
        help_text="Servidor IMAP (ex.: outlook.office365.com)"
    )
    port = models.IntegerField(
        default=993,
        help_text="Porta IMAP (normalmente 993 para IMAPS)"
    )
    username = models.CharField(
        max_length=255,
        help_text="Nome de utilizador / endereço de email"
    )
    password = models.CharField(
        max_length=255,
        help_text="Password ou token de aplicação (não usar password 'crua' em produção)"
    )
    folder = models.CharField(
        max_length=100,
        default="INBOX",
        help_text="Pasta IMAP a monitorizar (ex.: INBOX, Faturas, etc.)"
    )
    use_ssl = models.BooleanField(
        default=True,
        help_text="Usar ligação segura (SSL/TLS) para IMAP"
    )
    active = models.BooleanField(
        default=True,
        help_text="Se desactivar, a conta deixa de ser processada pelo AttachFlow"
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Conta de Email"
        verbose_name_plural = "Contas de Email"

    def __str__(self):
        return f"{self.name} ({self.username})"
