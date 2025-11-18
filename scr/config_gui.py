import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import configparser
import os
import importlib.util

# Caminho do arquivo de configuração
config_path = os.path.join("config", "config.ini")

# Verificar bibliotecas instaladas
has_pywin32 = importlib.util.find_spec("win32com") is not None
has_exchangelib = importlib.util.find_spec("exchangelib") is not None

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

# Criar janela principal
root = tk.Tk()
root.title("Configuração do AttachFlow")
root.geometry("400x320")

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
combo_conexao = ttk.Combobox(root, values=[])
combo_conexao.pack(pady=5)

# Lógica para definir opções de conexão
if has_pywin32 and has_exchangelib:
    combo_conexao["values"] = ["pywin32", "exchangelib"]
    combo_conexao.set(config_data.get('conexao', 'pywin32'))
elif has_pywin32:
    combo_conexao["values"] = ["pywin32"]
    combo_conexao.set("pywin32")
    combo_conexao.config(state="disabled")
elif has_exchangelib:
    combo_conexao["values"] = ["exchangelib"]
    combo_conexao.set("exchangelib")
    combo_conexao.config(state="disabled")
else:
    combo_conexao.set("Nenhuma biblioteca disponível")
    combo_conexao.config(state="disabled")
    messagebox.showerror("Erro", "Nenhuma biblioteca encontrada. Instale pywin32 ou exchangelib.")

label_extensao = tk.Label(root, text="Extensão (fixo): pdf")
label_extensao.pack(pady=10)

btn_salvar = tk.Button(root, text="Salvar Configuração", command=salvar_config)
btn_salvar.pack(pady=10)

# Desabilitar botão salvar se nenhuma biblioteca disponível
if not (has_pywin32 or has_exchangelib):
    btn_salvar.config(state="disabled")

root.mainloop()