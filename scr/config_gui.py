import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import configparser
import os
import importlib.util
import subprocess
import sys

# Caminho do arquivo de configuração
config_path = os.path.join("config", "config.ini")

# Função para verificar se uma biblioteca está instalada
def is_installed(module_name):
    return importlib.util.find_spec(module_name) is not None

# Função para instalar uma biblioteca via pip
def install_package(package_name):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package_name])
        return True
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao instalar {package_name}: {e}")
        return False

# Estado da validação
validacao_ok = False

# Função para carregar configurações existentes
def carregar_config():
    config = configparser.ConfigParser()
    if os.path.exists(config_path):
        config.read(config_path)
        return config['DEFAULT']
    else:
        return {'pasta_outlook': '', 'pasta_destino': '', 'extensao': 'pdf', 'conexao': ''}

# Função para salvar configurações
def salvar_config():
    if not validacao_ok:
        messagebox.showerror("Erro", "Valide a conexão antes de salvar.")
        return

    pasta_outlook = entry_outlook.get().strip()
    pasta_destino = entry_destino.get().strip()
    conexao = combo_conexao.get().strip()
    extensao = 'pdf'

    if not pasta_outlook or not pasta_destino or not conexao:
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios.")
        return

    config = configparser.ConfigParser()
    config['DEFAULT'] = {
        'pasta_outlook': pasta_outlook,
        'pasta_destino': pasta_destino,
        'extensao': extensao,
        'conexao': conexao
    }

    with open(config_path, 'w') as configfile:
        config.write(configfile)

    messagebox.showinfo("Sucesso", "Configuração salva com sucesso!")

# Função para escolher pasta destino
def escolher_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta destino")
    if pasta:
        entry_destino.delete(0, tk.END)
        entry_destino.insert(0, pasta)

# Função para validar conexão
def validar_conexao():
    global validacao_ok
    pasta_outlook = entry_outlook.get().strip()
    conexao = combo_conexao.get().strip()

    if not pasta_outlook or not conexao:
        messagebox.showerror("Erro", "Preencha os campos antes de validar.")
        return

    # Verifica se a biblioteca está instalada, se não estiver, oferece instalar
    if conexao == "pywin32" and not is_installed("win32com"):
        resposta = messagebox.askyesno("Instalar pywin32", "A biblioteca pywin32 não está instalada. Deseja instalar agora?")
        if resposta:
            if install_package("pywin32"):
                messagebox.showinfo("Instalado", "pywin32 instalada com sucesso. Tente validar novamente.")
            else:
                label_status.config(text="❌ Falha na instalação do pywin32.", fg="red")
            return
        else:
            label_status.config(text="❌ pywin32 não instalada.", fg="red")
            return

    if conexao == "exchangelib" and not is_installed("exchangelib"):
        resposta = messagebox.askyesno("Instalar exchangelib", "A biblioteca exchangelib não está instalada. Deseja instalar agora?")
        if resposta:
            if install_package("exchangelib"):
                messagebox.showinfo("Instalado", "exchangelib instalada com sucesso. Tente validar novamente.")
            else:
                label_status.config(text="❌ Falha na instalação do exchangelib.", fg="red")
            return
        else:
            label_status.config(text="❌ exchangelib não instalada.", fg="red")
            return

    try:
        if conexao == "pywin32":
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            pasta_encontrada = any(folder.Name == pasta_outlook for folder in inbox.Folders)
            if pasta_encontrada:
                label_status.config(text="✅ Conexão OK (pywin32)", fg="green")
                validacao_ok = True
            else:
                label_status.config(text="❌ Pasta não encontrada.", fg="red")
                validacao_ok = False

        elif conexao == "exchangelib":
            from exchangelib import Account, Credentials, DELEGATE
            email = simpledialog.askstring("Credenciais", "Informe seu e-mail:")
            senha = simpledialog.askstring("Credenciais", "Informe sua senha:", show='*')
            creds = Credentials(username=email, password=senha)
            try:
                account = Account(primary_smtp_address=email, credentials=creds, autodiscover=True, access_type=DELEGATE)
                pasta_encontrada = any(f.name == pasta_outlook for f in account.root.walk())
                if pasta_encontrada:
                    label_status.config(text="✅ Conexão OK (exchangelib)", fg="green")
                    validacao_ok = True
                else:
                    label_status.config(text="❌ Pasta não encontrada.", fg="red")
                    validacao_ok = False
            except Exception as e:
                label_status.config(text=f"❌ Erro: {e}", fg="red")
                validacao_ok = False

    except Exception as e:
        label_status.config(text=f"❌ Erro: {e}", fg="red")
        validacao_ok = False

# Criar janela principal
root = tk.Tk()
root.title("Configuração do AttachFlow")
root.geometry("400x400")

# Carregar configurações existentes
config_data = carregar_config()

# Labels e campos
label_outlook = tk.Label(root, text="Nome da pasta do Outlook:")
label_outlook.pack(pady=5)
entry_outlook = tk.Entry(root, width=40)
entry_outlook.pack(pady=5)
entry_outlook.insert(0, config_data.get('pasta_outlook', ''))

label_destino = tk.Label(root, text="Pasta destino:")
label_destino.pack(pady=5)
entry_destino = tk.Entry(root, width=40)
entry_destino.pack(pady=5)
entry_destino.insert(0, config_data.get('pasta_destino', ''))

btn_escolher = tk.Button(root, text="Escolher Pasta", command=escolher_pasta)
btn_escolher.pack(pady=5)

label_conexao = tk.Label(root, text="Método de conexão:")
label_conexao.pack(pady=5)
combo_conexao = ttk.Combobox(root, values=["pywin32", "exchangelib"])
combo_conexao.pack(pady=5)
combo_conexao.set(config_data.get('conexao', 'pywin32'))

label_extensao = tk.Label(root, text="Extensão (fixo): pdf")
label_extensao.pack(pady=10)

btn_validar = tk.Button(root, text="Validar Conexão", command=validar_conexao)
btn_validar.pack(pady=5)

label_status = tk.Label(root, text="Status: Não validado", fg="blue")
label_status.pack(pady=5)

btn_salvar = tk.Button(root, text="Salvar Configuração", command=salvar_config)
btn_salvar.pack(pady=10)

root.mainloop()