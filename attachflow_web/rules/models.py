from django.db import models
from accounts.models import EmailAccount

class Rule(models.Model):
    """
    Regra de captura de anexos para uma conta de email específica.
    Cada regra pode apontar para uma pasta IMAP diferente.    
    """
    account = models.ForeignKey(
        EmailAccount,
        on_delete=models.CASCADE,
        related_name="rules",
        help_text="Conta de email à qual esta regra pertence."
    )
    
    name = models.CharField(
        max_length=100,
        help_text="Nome amigável da regra (ex.: Faturas Continente, Relatórios Mensais, etc.)"
    )

    enabled = models.BooleanField(
        default=True,
        help_text="Se desactivada, a regra não será executada."
    )

    # Pasta IMAP específica para esta regra
    imap_folder = models.CharField(
        max_length=150,
        default="INBOX",
        help_text="Pasta IMAP onde esta regra será aplicada (ex.: INBOX, INBOX/Faturas, Faturas)."
    )

    # Filtros
    from_contains = models.CharField(
        max_length=255,
        blank=True,
        help_text="Filtro por remetente (contém). Ex.: faturas@supermercado.pt"
    )

    subject_contains = models.CharField(
        max_length=255,
        blank=True,
        help_text="Filtro por assunto (contém). Ex.: 'Fatura' ou 'Factura'"
    )

    filename_regex = models.CharField(
        max_length=255,
        blank=True,
        help_text="Expressão regular para o nome de ficheiro do anexo. Ex.: ^Fatura_.*\\.pdf$"
    )

    # Acções
    dest_folder = models.CharField(
        max_length=500,
        help_text="Pasta local para onde os anexos serão descarregados."
    )

    mark_as_read = models.BooleanField(
        default=True,
        help_text="Marcar email como lido após processamento."
    )

    move_to_folder = models.CharField(
        max_length=255,
        blank=True,
        help_text="Se preenchido, mover o email para esta pasta IMAP após processamento."
    )
    
    filename_template = models.CharField(
    max_length=255,
    default="{date:%Y.%m.%d %H.%M} - {rule_name}{index}{ext}",
    help_text=(
        "Template para o nome do ficheiro. "
        "Placeholders suportados: "
        "{date:FORMATO}, {rule_name}, {account_name}, "
        "{original_name}, {index}, {ext}."
        )
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

class ProcessedEmail(models.Model):
    """
    Marca emails que já foram processados por uma regra específica,
    para evitar downloads duplicados.
    """

    account = models.ForeignKey(
        EmailAccount,
        on_delete=models.CASCADE,
        related_name="processed_emails"
    )

    rule = models.ForeignKey(
        'Rule',
        on_delete=models.CASCADE,
        related_name="processed_emails"
    )

    # UID IMAP é mais fiável que só o ID interno da sessão
    folder = models.CharField(max_length=150)
    uid = models.CharField(max_length=100)

    message_id_header = models.CharField(
        max_length=255,
        blank=True,
        help_text="Valor do cabeçalho Message-ID (se disponível)."
    )

    processed_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('account', 'folder', 'uid', 'rule')
        verbose_name = "Email Processado"
        verbose_name_plural = "Emails Processados"

    def __str__(self):
        return f"{self.account} | {self.folder} | UID {self.uid} | Regra {self.rule_id}"
