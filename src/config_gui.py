"""Interface gráfica para configuração do AttachFlow.

O objetivo deste utilitário é facilitar a criação do arquivo config.ini
com todas as validações necessárias (dependências, conexão e caminhos).
"""
from __future__ import annotations

import configparser
import importlib.util
import os
import platform
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

CONFIG_PATH = Path("config/config.ini")
DEFAULTS = {
    "pasta_outlook": "",
    "pasta_destino": "",
    "extensao": "pdf",
    "conexao": "pywin32",
}
WINDOW_TITLE = "AttachFlow - Configuração"


@dataclass
class ConfigData:
    pasta_outlook: str
    pasta_destino: str
    extensao: str
    conexao: str

    @classmethod
    def from_parser(cls, data: Dict[str, str]) -> "ConfigData":
        return cls(
            pasta_outlook=data.get("pasta_outlook", ""),
            pasta_destino=data.get("pasta_destino", ""),
            extensao=data.get("extensao", "pdf"),
            conexao=data.get("conexao", "pywin32"),
        )

    def to_parser(self) -> configparser.ConfigParser:
        parser = configparser.ConfigParser()
        parser["DEFAULT"] = {
            "pasta_outlook": self.pasta_outlook.strip(),
            "pasta_destino": self.pasta_destino.strip(),
            "extensao": (self.extensao or "pdf").strip().lower(),
            "conexao": self.conexao.strip(),
        }
        return parser


class DependencyManager:
    @staticmethod
    def is_installed(module_name: str) -> bool:
        return importlib.util.find_spec(module_name) is not None

    @staticmethod
    def install(package_name: str) -> bool:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--upgrade", package_name]
            )
            return True
        except Exception as exc:  # pragma: no cover - interação com pip
            messagebox.showerror(
                "Erro",
                f"Falha ao instalar {package_name}.\n"
                f"Detalhes: {exc}\n\nTente executar manualmente: pip install {package_name}",
            )
            return False


class ConfigApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry("540x520")
        self.root.resizable(False, False)

        self.validacao_ok = False
        self.config_data = self._load_config()

        self.var_pasta_outlook = tk.StringVar(value=self.config_data.pasta_outlook)
        self.var_pasta_destino = tk.StringVar(value=self.config_data.pasta_destino)
        self.var_conexao = tk.StringVar(value=self.config_data.conexao)
        self.var_extensao = tk.StringVar(value=self.config_data.extensao)
        self.status_var = tk.StringVar(value="Status: aguardando validação...")

        self._build_layout()
        self._bind_events()

    # ------------------------------------------------------------------
    # Config helpers
    def _load_config(self) -> ConfigData:
        if CONFIG_PATH.exists():
            parser = configparser.ConfigParser()
            parser.read(CONFIG_PATH)
            return ConfigData.from_parser(parser["DEFAULT"])
        return ConfigData.from_parser(DEFAULTS)

    def _save_config(self) -> None:
        if not self.validacao_ok:
            raise RuntimeError("Valide a conexão antes de salvar.")
        config = ConfigData(
            pasta_outlook=self.var_pasta_outlook.get().strip(),
            pasta_destino=self.var_pasta_destino.get().strip(),
            extensao=self.var_extensao.get().strip() or "pdf",
            conexao=self.var_conexao.get().strip(),
        )
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with CONFIG_PATH.open("w") as config_file:
            config.to_parser().write(config_file)

    # ------------------------------------------------------------------
    # Layout
    def _build_layout(self) -> None:
        padding = {"padx": 15, "pady": 6}

        header = tk.Label(
            self.root,
            text=(
                "Configure abaixo os parâmetros utilizados pelo AttachFlow. "
                "Após validar a conexão, salve o arquivo config.ini."
            ),
            wraplength=500,
            justify="left",
        )
        header.pack(padx=20, pady=(15, 10), anchor="w")

        form_frame = ttk.LabelFrame(self.root, text="Parâmetros obrigatórios")
        form_frame.pack(fill="x", padx=15, pady=5)

        # Pasta outlook
        ttk.Label(form_frame, text="Nome da pasta no Outlook:").grid(row=0, column=0, sticky="w", **padding)
        outlook_container = ttk.Frame(form_frame)
        outlook_container.grid(row=0, column=1, sticky="we", **padding)
        outlook_container.columnconfigure(0, weight=1)
        self.entry_outlook = ttk.Entry(outlook_container, textvariable=self.var_pasta_outlook, width=36)
        self.entry_outlook.grid(row=0, column=0, sticky="we")
        ttk.Button(
            outlook_container,
            text="Ajuda",
            command=self._mostrar_dicas_pasta,
            width=10,
        ).grid(row=0, column=1, padx=(5, 0))

        # Pasta destino
        ttk.Label(form_frame, text="Pasta destino dos anexos:").grid(row=1, column=0, sticky="w", **padding)
        dest_container = ttk.Frame(form_frame)
        dest_container.grid(row=1, column=1, sticky="we", **padding)
        dest_container.columnconfigure(0, weight=1)
        self.entry_destino = ttk.Entry(dest_container, textvariable=self.var_pasta_destino, width=36)
        self.entry_destino.grid(row=0, column=0, sticky="we")
        ttk.Button(dest_container, text="Selecionar...", command=self._selecionar_pasta).grid(row=0, column=1, padx=(5, 0))

        # Método conexão
        ttk.Label(form_frame, text="Método de conexão:").grid(row=2, column=0, sticky="w", **padding)
        self.combo_conexao = ttk.Combobox(
            form_frame,
            textvariable=self.var_conexao,
            values=["pywin32", "exchangelib"],
            state="readonly",
            width=18,
        )
        self.combo_conexao.grid(row=2, column=1, sticky="w", **padding)

        # Extensão (somente leitura)
        ttk.Label(form_frame, text="Extensão monitorada:").grid(row=3, column=0, sticky="w", **padding)
        self.entry_extensao = ttk.Entry(form_frame, textvariable=self.var_extensao, state="readonly", width=15)
        self.entry_extensao.grid(row=3, column=1, sticky="w", **padding)

        # Informações adicionais
        info_box = ttk.LabelFrame(self.root, text="Dicas e pré-requisitos")
        info_box.pack(fill="both", padx=15, pady=10)
        info = (
            "• pywin32 está disponível apenas no Windows.\n"
            "• Para contas com MFA utilize senha de aplicativo (Exchange).\n"
            "• Certifique-se de ter permissão para criar arquivos na pasta destino."
        )
        ttk.Label(info_box, text=info, justify="left").pack(anchor="w", padx=10, pady=6)

        # Status + Log
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="both", padx=15, pady=(5, 10), expand=True)
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, foreground="#0f4c81")
        self.status_label.pack(anchor="w")
        self.log_box = tk.Text(status_frame, height=8, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True, pady=(5, 0))

        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=15, pady=10)
        ttk.Button(button_frame, text="Validar conexão", command=self._validar).pack(side="left")
        ttk.Button(button_frame, text="Salvar configurações", command=self._salvar).pack(side="right")

    def _bind_events(self) -> None:
        for widget in (self.entry_outlook, self.entry_destino):
            widget.bind("<KeyRelease>", self._invalidate_validation)
        self.combo_conexao.bind("<<ComboboxSelected>>", self._invalidate_validation)

    # ------------------------------------------------------------------
    # UI helpers
    def _selecionar_pasta(self) -> None:
        pasta = filedialog.askdirectory(title="Selecione a pasta destino")
        if pasta:
            self.var_pasta_destino.set(pasta)
            self._invalidate_validation()

    def _append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.configure(state="disabled")
        self.log_box.see("end")

    def _set_status(self, message: str, color: str = "#0f4c81") -> None:
        self.status_var.set(message)
        self.status_label.configure(foreground=color)
        self._append_log(message)

    def _invalidate_validation(self, *_args) -> None:
        if self.validacao_ok:
            self.validacao_ok = False
            self._set_status("Status: alterações detectadas. Valide novamente.", "#a15c0f")

    # ------------------------------------------------------------------
    # Validation / save
    def _normalize_folder_path(self, folder_path: str) -> List[str]:
        normalized = folder_path.replace("\\", "/").replace(">", "/")
        partes = [parte.strip() for parte in normalized.split("/") if parte.strip()]
        if not partes:
            raise ValueError("Informe o caminho da pasta no Outlook (use '/' para subpastas).")
        return partes

    def _check_destination(self) -> Path:
        destino_str = self.var_pasta_destino.get().strip()
        if not destino_str:
            raise ValueError("Informe a pasta destino.")
        destino = Path(os.path.expandvars(os.path.expanduser(destino_str)))
        if not destino.exists():
            criar = messagebox.askyesno(
                "Criar pasta",
                f"A pasta {destino} não existe. Deseja criá-la agora?",
            )
            if criar:
                destino.mkdir(parents=True, exist_ok=True)
            else:
                raise ValueError("Pasta destino inexistente.")
        if not os.access(destino, os.W_OK):
            raise PermissionError("Sem permissão de escrita na pasta destino.")
        return destino

    def _validar_dependencias(self, conexao: str) -> bool:
        if conexao == "pywin32" and platform.system().lower() != "windows":
            messagebox.showwarning(
                "Aviso",
                "pywin32 requer Windows. Altere o método de conexão ou execute em um ambiente compatível.",
            )
            return False

        module = "win32com" if conexao == "pywin32" else "exchangelib"
        if not DependencyManager.is_installed(module):
            resposta = messagebox.askyesno(
                "Dependência ausente",
                f"O módulo '{module}' não foi encontrado. Deseja instalar automaticamente?",
            )
            if resposta:
                pacote = "pywin32" if conexao == "pywin32" else "exchangelib"
                if not DependencyManager.install(pacote):
                    return False
                self._append_log(f"Biblioteca {pacote} instalada com sucesso.")
                return True
            return False
        return True

    def _validar(self) -> None:
        pasta_outlook = self.var_pasta_outlook.get().strip()
        conexao = self.var_conexao.get().strip()

        try:
            if not pasta_outlook:
                raise ValueError("Informe o nome da pasta a ser monitorada no Outlook.")
            self._check_destination()
            if not self._validar_dependencias(conexao):
                self._set_status("❌ Dependências não atendidas.", "red")
                return

            if conexao == "pywin32":
                self._validar_pywin32(pasta_outlook)
            else:
                self._validar_exchangelib(pasta_outlook)
        except Exception as exc:
            self.validacao_ok = False
            self._set_status(f"❌ Erro na validação: {exc}", "red")

    def _validar_pywin32(self, pasta_outlook: str) -> None:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 não está instalado. Utilize o botão validar novamente após a instalação."
            ) from exc

        partes = self._normalize_folder_path(pasta_outlook)
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        pasta_alvo = self._find_pywin32_folder(namespace, partes)
        if pasta_alvo is not None:
            self.validacao_ok = True
            self._set_status("✅ Conexão validada com sucesso via pywin32!", "green")
        else:
            raise RuntimeError("Pasta do Outlook não encontrada. Verifique o nome informado.")

    def _validar_exchangelib(self, pasta_outlook: str) -> None:
        from exchangelib import Account, Credentials, DELEGATE

        email = simpledialog.askstring("Credenciais", "Informe seu e-mail:")
        senha = simpledialog.askstring("Credenciais", "Informe sua senha:", show="*")
        if not email or not senha:
            raise RuntimeError("Credenciais não fornecidas.")

        creds = Credentials(username=email, password=senha)
        account = Account(primary_smtp_address=email, credentials=creds, autodiscover=True, access_type=DELEGATE)
        partes = self._normalize_folder_path(pasta_outlook)
        pasta_encontrada = self._find_exchangelib_folder(account.root, partes)
        if pasta_encontrada:
            self.validacao_ok = True
            self._set_status("✅ Conexão validada com sucesso via exchangelib!", "green")
        else:
            raise RuntimeError("Pasta não encontrada na conta Exchange informada.")

    def _salvar(self) -> None:
        try:
            self._save_config()
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            return
        messagebox.showinfo("Sucesso", f"Configurações salvas em {CONFIG_PATH}!")
        self._append_log("Arquivo config.ini atualizado com sucesso.")

    # ------------------------------------------------------------------
    # Outlook helpers
    def _mostrar_dicas_pasta(self) -> None:
        dicas_base = [
            "• Utilize '/' ou '>' para separar subpastas (ex.: Caixa de Entrada/Financeiro/2024)",
            "• Para pastas compartilhadas, informe o caminho completo partindo da raiz",
        ]
        conexao = self.var_conexao.get()
        try:
            if conexao == "pywin32" and DependencyManager.is_installed("win32com"):
                nomes = self._listar_pastas_pywin32()
                if nomes:
                    lista = "\n- ".join(nomes)
                    messagebox.showinfo(
                        "Pastas detectadas",
                        "Algumas pastas detectadas via Outlook:\n- " + lista,
                    )
                    return
        except Exception as exc:  # pragma: no cover - interação com Outlook
            self._append_log(f"Não foi possível listar pastas: {exc}")
        messagebox.showinfo("Ajuda", "\n".join(dicas_base))

    def _listar_pastas_pywin32(self) -> List[str]:
        import win32com.client  # type: ignore

        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        encontrados: List[str] = []

        def walk(folder, prefix: str) -> None:
            if len(encontrados) >= 20:
                return
            for child in folder.Folders:
                caminho = f"{prefix}/{child.Name}"
                encontrados.append(caminho)
                walk(child, caminho)

        walk(inbox, getattr(inbox, "Name", "Caixa de Entrada"))
        return encontrados

    def _find_pywin32_folder(self, namespace, partes: List[str]):
        inbox = namespace.GetDefaultFolder(6)

        def descend(folder, remaining: List[str]):
            if not remaining:
                return folder
            alvo = remaining[0].lower()
            for child in folder.Folders:
                if child.Name.lower() == alvo:
                    return descend(child, remaining[1:])
            return None

        if partes[0].lower() == (getattr(inbox, "Name", "").lower()):
            return descend(inbox, partes[1:])

        for root_folder in namespace.Folders:
            if root_folder.Name.lower() == partes[0].lower():
                return descend(root_folder, partes[1:])

        return descend(inbox, partes)

    def _find_exchangelib_folder(self, folder, partes: List[str]):
        if not partes:
            return folder
        alvo = partes[0].lower()
        for child in folder.children:
            nome = getattr(child, "name", "")
            if nome and nome.lower() == alvo:
                return self._find_exchangelib_folder(child, partes[1:])
        return None

    # ------------------------------------------------------------------
    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = ConfigApp()
    app.run()


if __name__ == "__main__":
    main()
