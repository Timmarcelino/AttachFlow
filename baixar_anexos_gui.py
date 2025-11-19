"""Interface gráfica principal do AttachFlow.

Este script lê as configurações salvas pelo config_gui.py e executa o download
 dos anexos utilizando o método de conexão selecionado (pywin32 ou exchangelib).
"""
from __future__ import annotations

import configparser
import datetime as dt
import logging
import queue
import re
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk, simpledialog
from typing import List

import pandas as pd

CONFIG_FILE = Path("config/config.ini")
LOG_FILE = Path("logs/download_anexos.log")
REPORT_FILE = Path("reports/relatorio_anexos.xlsx")


@dataclass
class AppConfig:
    pasta_outlook: str
    pasta_destino: Path
    conexao: str
    extensao: str = "pdf"

    @classmethod
    def load(cls) -> "AppConfig":
        if not CONFIG_FILE.exists():
            raise FileNotFoundError(
                f"Arquivo de configuração não encontrado em {CONFIG_FILE.resolve()}. "
                "Execute o config_gui.py antes de iniciar o download."
            )
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE)
        data = config["DEFAULT"]
        pasta_destino = Path(data.get("pasta_destino", "").strip())
        if not pasta_destino:
            raise ValueError("Pasta destino não configurada.")
        conexao = data.get("conexao", "pywin32").lower().strip()
        if conexao not in {"pywin32", "exchangelib"}:
            raise ValueError("Método de conexão inválido no config.ini.")
        return cls(
            pasta_outlook=data.get("pasta_outlook", "").strip(),
            pasta_destino=pasta_destino,
            conexao=conexao,
            extensao=data.get("extensao", "pdf").strip().lower() or "pdf",
        )


class AttachmentDownloader:
    def __init__(self, config: AppConfig):
        self.config = config
        self.records: List[dict] = []
        self.total_processed = 0
        self.total_downloaded = 0

    def run(self, progress_queue: "queue.Queue[str]") -> None:
        logging.info("Iniciando rotina de download - método: %s", self.config.conexao)
        self.config.pasta_destino.mkdir(parents=True, exist_ok=True)
        if self.config.conexao == "pywin32":
            self._process_pywin32(progress_queue)
        else:
            self._process_exchangelib(progress_queue)
        self._generate_report()
        logging.info(
            "Processo finalizado. Total processado: %s | Total salvo: %s",
            self.total_processed,
            self.total_downloaded,
        )

    def _generate_report(self) -> None:
        if not self.records:
            logging.info("Nenhum anexo baixado. Relatório não será gerado.")
            return
        REPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
        df = pd.DataFrame(self.records)
        df.sort_values(by="data_email", inplace=True)
        df.to_excel(REPORT_FILE, index=False)
        logging.info("Relatório salvo em %s", REPORT_FILE.resolve())

    def _sanitize_sender(self, sender: str) -> str:
        sender = sender or "Remetente desconhecido"
        sender = re.sub(r"[\\/:*?\"<>|]", " ", sender)
        sender = re.sub(r"\s+", " ", sender).strip()
        return sender or "Remetente"

    def _build_filename(self, timestamp: dt.datetime, sender: str) -> str:
        formatted_date = timestamp.strftime("%Y.%m.%d %H.%M")
        safe_sender = self._sanitize_sender(sender)
        return f"{formatted_date} - {safe_sender}.pdf"

    def _save_record(
        self,
        message_date: dt.datetime,
        sender: str,
        subject: str,
        saved_path: Path,
    ) -> None:
        self.records.append(
            {
                "data_email": message_date,
                "remetente": sender,
                "assunto": subject,
                "arquivo": saved_path.name,
                "caminho": str(saved_path),
            }
        )

    def _process_pywin32(self, progress_queue: "queue.Queue[str]") -> None:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError(
                "pywin32 não está instalado. Execute a validação pelo config_gui.py."
            ) from exc

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folder = self._find_pywin32_folder(outlook, self.config.pasta_outlook)
        if folder is None:
            raise RuntimeError(
                f"Pasta '{self.config.pasta_outlook}' não encontrada no Outlook."
            )

        items = folder.Items
        items = items.Restrict("[HasAttachments] = True")
        count = items.Count
        logging.info("%s mensagens com anexos encontradas via pywin32.", count)
        for message in items:
            self.total_processed += 1
            progress_queue.put(f"Processando e-mail {self.total_processed}/{count}")
            sender = getattr(message, "SenderName", "Desconhecido")
            subject = getattr(message, "Subject", "")
            message_date = getattr(message, "SentOn", None) or getattr(
                message, "ReceivedTime", dt.datetime.now()
            )
            for attachment in getattr(message, "Attachments", []):
                filename = getattr(attachment, "FileName", "")
                if not filename.lower().endswith(f".{self.config.extensao}"):
                    continue
                target_name = self._build_filename(message_date, sender)
                target_path = self.config.pasta_destino / target_name
                if target_path.exists():
                    logging.info("Arquivo já existe, ignorado: %s", target_path)
                    continue
                try:
                    attachment.SaveAsFile(str(target_path))
                    self.total_downloaded += 1
                    logging.info("Anexo salvo: %s", target_path)
                    self._save_record(message_date, sender, subject, target_path)
                except Exception as exc:  # pragma: no cover
                    logging.exception("Falha ao salvar anexo: %s", exc)

    def _process_exchangelib(self, progress_queue: "queue.Queue[str]") -> None:
        try:
            from exchangelib import Account, Credentials, DELEGATE
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError(
                "exchangelib não está instalado. Execute a validação pelo config_gui.py."
            ) from exc

        email = simpledialog.askstring("Credenciais", "Informe seu e-mail:")
        senha = simpledialog.askstring("Credenciais", "Informe sua senha:", show="*")
        if not email or not senha:
            raise RuntimeError("Credenciais não informadas para conexão exchangelib.")

        creds = Credentials(username=email, password=senha)
        account = Account(
            primary_smtp_address=email,
            credentials=creds,
            autodiscover=True,
            access_type=DELEGATE,
        )
        folder = self._find_exchangelib_folder(account, self.config.pasta_outlook)
        if folder is None:
            raise RuntimeError(
                f"Pasta '{self.config.pasta_outlook}' não encontrada na conta Exchange."
            )

        queryset = folder.filter(has_attachments=True)
        count = queryset.count() if hasattr(queryset, "count") else None
        logging.info("Processando anexos via exchangelib. Total estimado: %s", count)
        for idx, message in enumerate(queryset, start=1):
            self.total_processed += 1
            progress_label = (
                f"Processando e-mail {idx}/{count}" if count else f"Processando e-mail {idx}"
            )
            progress_queue.put(progress_label)
            sender = getattr(message, "sender", None)
            sender_name = getattr(sender, "name", None) if sender else "Desconhecido"
            subject = getattr(message, "subject", "")
            message_date = getattr(message, "datetime_sent", None) or getattr(
                message, "datetime_received", dt.datetime.now()
            )
            for attachment in getattr(message, "attachments", []):
                if not getattr(attachment, "is_inline", False) and getattr(
                    attachment, "name", ""
                ).lower().endswith(f".{self.config.extensao}"):
                    target_name = self._build_filename(message_date, sender_name)
                    target_path = self.config.pasta_destino / target_name
                    if target_path.exists():
                        logging.info("Arquivo já existe, ignorado: %s", target_path)
                        continue
                    with open(target_path, "wb") as file_handle:
                        file_handle.write(attachment.content)
                    self.total_downloaded += 1
                    logging.info("Anexo salvo: %s", target_path)
                    self._save_record(message_date, sender_name, subject, target_path)

    def _find_pywin32_folder(self, namespace, folder_name: str):
        if not folder_name:
            return namespace.GetDefaultFolder(6)  # Inbox por padrão

        def walk(folder):
            if folder.Name == folder_name:
                return folder
            for sub in folder.Folders:
                found = walk(sub)
                if found is not None:
                    return found
            return None

        # Procura na raiz padrão e depois em todas as stores
        inbox = namespace.GetDefaultFolder(6)
        found = walk(inbox)
        if found is not None:
            return found
        for store in namespace.Folders:
            found = walk(store)
            if found is not None:
                return found
        return None

    def _find_exchangelib_folder(self, account, folder_name: str):
        if not folder_name:
            return account.inbox
        for folder in account.root.walk():
            if folder.name == folder_name:
                return folder
        return None


def setup_logging() -> None:
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def main() -> None:
    setup_logging()
    root = tk.Tk()
    root.title("AttachFlow - Download de anexos")
    root.geometry("460x220")

    status_var = tk.StringVar(value="Aguardando início...")

    progress = ttk.Progressbar(root, mode="indeterminate")
    progress.pack(fill="x", padx=20, pady=20)

    status_label = tk.Label(root, textvariable=status_var)
    status_label.pack(pady=5)

    progress_queue: "queue.Queue[str]" = queue.Queue()

    def update_status() -> None:
        try:
            msg = progress_queue.get_nowait()
            status_var.set(msg)
        except queue.Empty:
            pass
        finally:
            root.after(300, update_status)

    def run_download() -> None:
        progress.start(10)
        try:
            config = AppConfig.load()
            downloader = AttachmentDownloader(config)
            downloader.run(progress_queue)
            messagebox.showinfo(
                "Concluído",
                f"Processo finalizado. {downloader.total_downloaded} anexos salvos.",
            )
            status_var.set(
                f"Concluído: {downloader.total_downloaded} anexos salvos."
            )
        except Exception as exc:
            logging.exception("Erro durante execução: %s", exc)
            messagebox.showerror("Erro", str(exc))
            status_var.set(f"Erro: {exc}")
        finally:
            progress.stop()

    def on_start():
        threading.Thread(target=run_download, daemon=True).start()

    start_button = tk.Button(root, text="Iniciar Download", command=on_start)
    start_button.pack(pady=10)

    update_status()
    root.mainloop()


if __name__ == "__main__":
    main()
