import tkinter as tk
from tkinter import filedialog, messagebox
import configparser
import os

# Caminho do arquivo de configuração
config_path = os.path.join("config", "config.ini")

# Função para carregar configurações existentes
def carregar_config():
    config = configparser.ConfigParser()
    if os.path.exists(config_path):
        config.read(config_path)
        return config['DEFAULT']
    else:
        return {'pasta_outlook': '', 'pasta_destino': '', 'extensao': 'pdf'}

# Função para salvar configurações
def salvar_config():
    pasta_outlook = entry_outlook.get().strip()
    pasta_destino = entry_destino.get().strip()
    extensao = 'pdf'  # fixo

    if not pasta_outlook or not pasta_destino:
        messagebox.showerror("Erro", "Preencha todos os campos obrigatórios.")
        return

    config = configparser.ConfigParser()
    config['DEFAULT'] = {
        'pasta_outlook': pasta_outlook,
        'pasta_destino': pasta_destino,
        'extensao': extensao
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
root.geometry("400x250")

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

label_extensao = tk.Label(root, text="Extensão (fixo): pdf")
label_extensao.pack(pady=10)

btn_salvar = tk.Button(root, text="Salvar Configuração", command=salvar_config)
btn_salvar.pack(pady=10)

root.mainloop()