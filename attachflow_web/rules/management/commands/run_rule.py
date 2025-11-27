from django.core.management.base import BaseCommand, CommandError
from rules.models import Rule
from accounts.models import EmailAccount

# Importar núcleo
from attachflow_core.email_core import (
    EmailAccountConfig,
    RuleConfig,
    run_rule
)

class Command(BaseCommand):
    help = "Executa uma regra específica do AttachFlow (por ID)."
    
    def add_arguments(self, parser):
        parser.add_argument("rule_id", type=int, help="ID da regra (Rule.id)")
        
    def handle(self, *args, **options):
        rule_id = options["rule_id"]
        
        try:
            rule = Rule.objects.select_related("account").get(id=rule_id)
        except Rule.DoesNotExist:
            raise CommandError(f"Regra com ID {rule_id} não encontrada.")
        
        account = rule.account
        
        # Construir EmailAccountConfig (núcleo)
        account_cfg = EmailAccountConfig(
            host=account.host,
            port=account.port,
            username=account.username,
            password=account.password,
            use_ssl=account.use_ssl,
            folder=account.folder,
            name=account.name,
        )
        
        # Construir RuleConfig (núcleo)
        rule_cfg = RuleConfig(
            account=account_cfg,
            name=rule.name,
            imap_folder=rule.imap_folder,
            from_contains=rule.from_contains,
            subject_contains=rule.subject_contains,
            filename_regex=rule.filename_regex,
            dest_folder=rule.dest_folder,
            mark_as_read=rule.mark_as_read,
            move_to_folder=rule.move_to_folder,
            filename_template=rule.filename_template,
        )
        
        self.stdout.write(self.style.NOTICE(f"▶ A executar regra: {rule.name}"))
        
        # Executar núcleo
        stats = run_rule(rule_cfg)
        
        # Mostrar resultados (MVP)
        self.stdout.write(self.style.SUCCESS(
            f"✔ Execução concluída.\n"
            f"Emails processados: {stats.emails_processed}\n"
            f"Anexos descarregados: {stats.attachments_downloaded}\n"
        ))
        
        if stats.errors:
            self.stdout.write(self.style.WARNING("⚠ Erros encontrados:"))
            for err in stats.errors:
                self.stdout.write(f" - {err}")
                
