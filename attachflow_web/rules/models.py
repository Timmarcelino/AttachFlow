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
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Regra de Captura"
        verbose_name_plural = "Regras de Captura"

    def __str__(self):
        return f"[{self.account.name}] {self.name}"
    
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
        Rule,
        on_delete=models.CASCADE,
        related_name="processed_emails"
    )

    folder = models.CharField(
        max_length=150,
        help_text="Pasta IMAP em que o email foi encontrado."
    )

    uid = models.CharField(
        max_length=100,
        help_text="UID IMAP da mensagem."
    )

    message_id_header = models.CharField(
        max_length=255,
        blank=True,
        help_text="Valor do cabeçalho Message-ID (se disponível)."
    )

    processed_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Email Processado"
        verbose_name_plural = "Emails Processados"
        unique_together = ('account', 'folder', 'uid', 'rule')

    def __str__(self):
        return f"{self.account} | {self.folder} | UID {self.uid} | Regra {self.rule_id}"

class JobExecution(models.Model):
    """
    Registo de uma execução de regra (run_rule).
    Pode representar uma execução manual ou agendada.
    """

    STATUS_CHOICES = [
        ("RUNNING", "Em execução"),
        ("SUCCESS", "Sucesso"),
        ("ERROR", "Erro"),
    ]

    rule = models.ForeignKey(
        Rule,
        on_delete=models.CASCADE,
        related_name="executions"
    )

    started_at = models.DateTimeField(auto_now_add=True)
    finished_at = models.DateTimeField(null=True, blank=True)

    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default="RUNNING"
    )

    emails_processed = models.IntegerField(default=0)
    attachments_downloaded = models.IntegerField(default=0)

    error_message = models.TextField(
        blank=True,
        help_text="Mensagem de erro, se a execução terminar em erro."
    )

    class Meta:
        verbose_name = "Execução de Regra"
        verbose_name_plural = "Execuções de Regras"
        ordering = ['-started_at']

    def __str__(self):
        return f"{self.rule.name} @ {self.started_at:%Y-%m-%d %H:%M}"

class AttachmentLog(models.Model):
    """
    Registo de cada anexo descarregado numa determinada execução.
    Facilita auditoria e estatísticas.
    """

    execution = models.ForeignKey(
        JobExecution,
        on_delete=models.CASCADE,
        related_name="attachments"
    )

    rule = models.ForeignKey(
        Rule,
        on_delete=models.CASCADE,
        related_name="attachment_logs"
    )

    account = models.ForeignKey(
        EmailAccount,
        on_delete=models.CASCADE,
        related_name="attachment_logs"
    )

    file_path = models.CharField(
        max_length=500,
        help_text="Caminho final do ficheiro no sistema de ficheiros."
    )

    file_name = models.CharField(
        max_length=255,
        help_text="Nome final do ficheiro (sem caminho)."
    )

    original_filename = models.CharField(
        max_length=255,
        help_text="Nome original do anexo no email."
    )

    size_bytes = models.BigIntegerField(default=0)

    sender = models.CharField(
        max_length=255,
        blank=True,
        help_text="Remetente (From) do email."
    )

    subject = models.CharField(
        max_length=500,
        blank=True,
        help_text="Assunto do email."
    )

    folder = models.CharField(
        max_length=150,
        help_text="Pasta IMAP em que o email foi encontrado."
    )

    uid = models.CharField(
        max_length=100,
        help_text="UID IMAP da mensagem."
    )

    message_id = models.CharField(
        max_length=255,
        blank=True,
        help_text="Cabeçalho Message-ID da mensagem."
    )

    downloaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Anexo Descarregado"
        verbose_name_plural = "Anexos Descarregados"
        ordering = ['-downloaded_at']

    def __str__(self):
        return self.file_path
