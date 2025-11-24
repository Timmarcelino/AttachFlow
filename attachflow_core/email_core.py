"""
Módulo core do AttachFlow para:
- Testar ligação IMAP
- Listar pastas IMAP
- Executar regras de download de anexos

Este módulo é independente do Django. O Django passa configurações simples
(dicionários / dataclasses) e recebe de volta resultados/estatísticas.
"""

from __future__ import annotations

import email
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from email.message import Message
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple

from imapclient import IMAPClient


# ------------------------------------------------------------------------------
# Logging básico
# ------------------------------------------------------------------------------

log = logging.getLogger(__name__)
if not log.handlers:
    # Configuração simples; em produção o Django pode sobrepor isto
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )


# ------------------------------------------------------------------------------
# Dataclasses de configuração e resultados
# ------------------------------------------------------------------------------

@dataclass
class EmailAccountConfig:
    host: str
    port: int
    username: str
    password: str
    use_ssl: bool = True
    folder: str = "INBOX"
    name: str = ""  # Nome amigável (opcional, para logs)


@dataclass
class RuleConfig:
    account: EmailAccountConfig
    name: str
    imap_folder: str = "INBOX"
    from_contains: str = ""
    subject_contains: str = ""
    filename_regex: str = ""
    dest_folder: str = ""
    mark_as_read: bool = True
    move_to_folder: str = ""
    filename_template: str = "{date:%Y.%m.%d %H.%M} - {rule_name}{index}{ext}"


@dataclass
class AttachmentDownload:
    rule_name: str
    account_name: str
    folder: str
    uid: str
    message_id: str
    original_filename: str
    final_path: Path
    size_bytes: int


@dataclass
class RunStats:
    emails_processed: int = 0
    attachments_downloaded: int = 0
    errors: List[str] = field(default_factory=list)
    downloaded_attachments: List[AttachmentDownload] = field(default_factory=list)
    processed_uids: Set[str] = field(default_factory=set)


@dataclass
class TestConnectionResult:
    success: bool
    message: str
    folders: List[str]


# ------------------------------------------------------------------------------
# Funções utilitárias
# ------------------------------------------------------------------------------

def _connect_imap(cfg: EmailAccountConfig) -> IMAPClient:
    """
    Cria e devolve um cliente IMAP ligado e autenticado.
    """
    log.debug("A ligar ao IMAP %s:%s (SSL=%s)", cfg.host, cfg.port, cfg.use_ssl)
    client = IMAPClient(cfg.host, port=cfg.port, ssl=cfg.use_ssl)
    client.login(cfg.username, cfg.password)
    return client


def _get_message_datetime(msg: Message) -> Optional[datetime]:
    """
    Tenta obter a data do email a partir do cabeçalho 'Date'.
    Se falhar, devolve None.
    """
    date_header = msg.get("Date")
    if not date_header:
        return None
    try:
        from email.utils import parsedate_to_datetime
        return parsedate_to_datetime(date_header)
    except Exception:
        return None


def _get_message_id(msg: Message) -> str:
    """
    Devolve o cabeçalho Message-ID se existir.
    """
    return msg.get("Message-ID", "")


def _format_date_in_template(template_str: str, dt: datetime) -> str:
    """
    Substitui ocorrências de {date:%Y-%m-%d ...} pelo strftime correspondente.
    """
    def repl(match):
        fmt = match.group(1)
        return dt.strftime(fmt)

    return re.sub(r"{date:([^}]+)}", repl, template_str)


def _sanitize_filename(name: str) -> str:
    """
    Remove caracteres proibidos para ficheiros em Windows/Linux.
    """
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def build_download_filename(
    msg: Message,
    attachment_name: str,
    rule: RuleConfig,
    index: int,
    dest_folder: Path,
) -> Path:
    """
    Gera o nome final do ficheiro, garantindo:
    - {index} vem antes de {ext}
    - unicidade dentro da pasta de destino
    - template por defeito:
      {date:%Y.%m.%d %H.%M} - {rule_name}{index}{ext}

    index:
      - 1  -> não acrescenta nada (ficheiro "limpo")
      - >1 -> acrescenta "_2", "_3", etc., conforme template
    """
    msg_date = _get_message_datetime(msg) or datetime.now()

    original = Path(attachment_name)
    ext = original.suffix or ""
    original_stem = original.stem

    template = rule.filename_template or "{date:%Y.%m.%d %H.%M} - {rule_name}{index}{ext}"

    # 1) aplicar {date:...}
    filled = _format_date_in_template(template, msg_date)

    # 2) aplicar os restantes placeholders
    # index: para o primeiro anexo não queremos "_1" por defeito
    index_placeholder = f"_{index}" if index > 1 else ""

    filled = filled.format(
        rule_name=rule.name,
        account_name=rule.account.name or rule.account.username,
        original_name=original_stem,
        index=index_placeholder,
        ext=ext,
    )

    safe_name = _sanitize_filename(filled)
    final_path = dest_folder / safe_name

    # 3) fallback se já existir: acrescentar sufixo _2, _3, ...
    counter = 2
    while final_path.exists():
        base = safe_name[:-len(ext)] if ext else safe_name
        candidate = f"{base}_{counter}{ext}"
        final_path = dest_folder / candidate
        counter += 1

    return final_path


# ------------------------------------------------------------------------------
# Teste de ligação e listagem de pastas
# ------------------------------------------------------------------------------

def test_imap_connection(cfg: EmailAccountConfig) -> TestConnectionResult:
    """
    Tenta ligar ao servidor IMAP, autenticar e listar pastas.
    Devolve sucesso/erro + lista de pastas (caso sucesso).
    """
    try:
        client = _connect_imap(cfg)
        folders_info = client.list_folders()
        # folders_info: [(flags, delimiter, name), ...]
        folders = [f[2] for f in folders_info]
        client.logout()
        return TestConnectionResult(
            success=True,
            message="Ligação bem sucedida.",
            folders=folders,
        )
    except Exception as e:
        log.exception("Erro ao testar ligação IMAP")
        return TestConnectionResult(
            success=False,
            message=f"Erro na ligação: {e}",
            folders=[],
        )


def list_imap_folders(cfg: EmailAccountConfig) -> List[str]:
    """
    Apenas lista as pastas IMAP. Assumimos que a ligação é válida.
    """
    client = _connect_imap(cfg)
    folders_info = client.list_folders()
    folders = [f[2] for f in folders_info]
    client.logout()
    return folders


# ------------------------------------------------------------------------------
# Execução de regras (run_rule)
# ------------------------------------------------------------------------------

def _build_search_criteria(rule: RuleConfig) -> List:
    """
    Constrói a lista de critérios IMAP para o IMAPClient.
    Só usamos filtros simples (FROM / SUBJECT); datas podem ser adicionadas depois.
    """
    criteria: List = ["ALL"]

    if rule.from_contains:
        criteria += ["FROM", rule.from_contains]

    if rule.subject_contains:
        criteria += ["SUBJECT", rule.subject_contains]

    return criteria


def _iter_attachments(msg: Message) -> Iterable[Tuple[str, bytes]]:
    """
    Itera sobre as partes da mensagem que são anexos.
    Devolve tuplos (filename, payload_bytes).
    """
    for part in msg.walk():
        content_disposition = part.get_content_disposition()
        if content_disposition != "attachment":
            continue

        filename = part.get_filename()
        if not filename:
            continue

        payload = part.get_payload(decode=True) or b""
        yield filename, payload


def run_rule(
    rule: RuleConfig,
    processed_uids: Optional[Set[str]] = None,
) -> RunStats:
    """
    Executa uma regra:
    - liga à conta IMAP associada
    - selecciona a pasta imap_folder
    - pesquisa mensagens com base em from/subject
    - descarrega anexos que obedeçam ao filename_regex
    - renomeia ficheiros segundo o filename_template
    - devolve RunStats com info de anexos descarregados e UIDs processados

    processed_uids:
      conjunto de UIDs que já foram marcados como processados (opcional).
      A função devolve os novos UIDs processados em RunStats.processed_uids
      para que o chamador (Django) possa persistir isso.
    """
    if processed_uids is None:
        processed_uids = set()

    stats = RunStats()

    dest_folder = Path(rule.dest_folder).expanduser()
    dest_folder.mkdir(parents=True, exist_ok=True)

    account_cfg = rule.account

    log.info(
        "A executar regra '%s' na conta '%s', pasta '%s'",
        rule.name,
        account_cfg.name or account_cfg.username,
        rule.imap_folder,
    )

    try:
        client = _connect_imap(account_cfg)
    except Exception as e:
        msg = f"Erro ao ligar à conta IMAP: {e}"
        log.error(msg)
        stats.errors.append(msg)
        return stats

    try:
        # seleccionar pasta IMAP
        client.select_folder(rule.imap_folder, readonly=False)

        # construir critérios de pesquisa
        criteria = _build_search_criteria(rule)
        log.debug("Critérios IMAP: %s", criteria)

        # search devolve UIDs por defeito no IMAPClient
        uids = client.search(criteria)
        log.info("Encontrados %d emails para a regra '%s'", len(uids), rule.name)

        for uid in uids:
            uid_str = str(uid)

            # já processado?
            if uid_str in processed_uids:
                log.debug("UID %s já processado, a ignorar.", uid_str)
                continue

            # fetch mensagem
            try:
                fetch_data = client.fetch([uid], ["RFC822"])
            except Exception as e:
                err = f"Erro ao buscar mensagem UID {uid}: {e}"
                log.error(err)
                stats.errors.append(err)
                continue

            raw_msg = fetch_data[uid][b"RFC822"]
            msg = email.message_from_bytes(raw_msg)

            stats.emails_processed += 1

            message_id = _get_message_id(msg)
            attachment_index = 0

            for filename, payload in _iter_attachments(msg):
                # aplicar regex ao nome do ficheiro, se definido
                if rule.filename_regex:
                    if not re.search(rule.filename_regex, filename):
                        log.debug(
                            "Anexo '%s' não corresponde ao padrão '%s', a ignorar.",
                            filename,
                            rule.filename_regex,
                        )
                        continue

                attachment_index += 1

                final_path = build_download_filename(
                    msg=msg,
                    attachment_name=filename,
                    rule=rule,
                    index=attachment_index,
                    dest_folder=dest_folder,
                )

                try:
                    with final_path.open("wb") as f:
                        f.write(payload)
                except Exception as e:
                    err = f"Erro ao gravar anexo '{final_path}': {e}"
                    log.error(err)
                    stats.errors.append(err)
                    continue

                size_bytes = len(payload)
                stats.attachments_downloaded += 1
                stats.downloaded_attachments.append(
                    AttachmentDownload(
                        rule_name=rule.name,
                        account_name=account_cfg.name or account_cfg.username,
                        folder=rule.imap_folder,
                        uid=uid_str,
                        message_id=message_id,
                        original_filename=filename,
                        final_path=final_path,
                        size_bytes=size_bytes,
                    )
                )

            # se algum anexo foi processado, marcamos o UID como processado
            if attachment_index > 0:
                stats.processed_uids.add(uid_str)

                # marcar como lido
                if rule.mark_as_read:
                    try:
                        client.add_flags([uid], [b"\\Seen"])
                    except Exception as e:
                        log.warning(
                            "Não foi possível marcar email UID %s como lido: %s",
                            uid_str,
                            e,
                        )

                # mover para outra pasta, se configurado
                if rule.move_to_folder:
                    try:
                        client.move([uid], rule.move_to_folder)
                    except Exception as e:
                        log.warning(
                            "Não foi possível mover email UID %s para '%s': %s",
                            uid_str,
                            rule.move_to_folder,
                            e,
                        )

    finally:
        try:
            client.logout()
        except Exception:
            pass

    log.info(
        "Regra '%s' concluída: %d emails processados, %d anexos descarregados.",
        rule.name,
        stats.emails_processed,
        stats.attachments_downloaded,
    )
    return stats
