# 📌 AttachFlow

Automação para download e organização de anexos PDF do Outlook com interface gráfica, renomeação inteligente e geração de relatórios.

---

## ✅ Descrição
O **AttachFlow** é uma ferramenta que conecta ao Microsoft Outlook, acessa uma pasta específica de e-mails e baixa anexos PDF automaticamente.  
Os arquivos são renomeados seguindo o padrão:

```

YYYY.MM.DD HH.MM - Remetente.pdf

````

Além disso, o sistema:
- Evita duplicidades (verifica se o arquivo já existe).
- Permite escolher a pasta destino via interface gráfica.
- Exibe barra de progresso durante o processamento.
- Gera **log detalhado** e um **relatório Excel** com todos os anexos baixados.

---

## 🚀 Funcionalidades
- Conexão automática ao Outlook via API nativa.
- Interface gráfica simples e intuitiva.
- Renomeação inteligente com data, hora e remetente.
- Filtro para anexos **PDF**.
- Geração de relatório em **Excel**.
- Log de execução para auditoria.
- Preparado para agendamento diário.

---

## 🛠 Tecnologias
- **Python 3.x**
- **pywin32** (integração com Outlook)
- **tkinter** (interface gráfica)
- **pandas + openpyxl** (relatório Excel)
- **logging** (logs detalhados)

---

## 📦 Instalação
1. Clone o repositório:
   ```bash
   git clone https://github.com/seuusuario/AttachFlow.git
   cd AttachFlow
````

2.  Crie um ambiente virtual:
    ```bash
    python -m venv venv
    source venv/bin/activate   # Linux/Mac
    venv\\Scripts\\activate    # Windows
    ```
3.  Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

***

## ▶ Como usar

1.  Execute o script principal:
    ```bash
    python baixar_anexos_gui.py
    ```
2.  Escolha:
    *   **Pasta do Outlook** (onde estão os e-mails filtrados).
    *   **Pasta destino** para salvar anexos.
3.  Clique em **Iniciar Download**.
4.  Ao final:
    *   Arquivos PDF renomeados estarão na pasta escolhida.
    *   Log gerado em `download_anexos.log`.
    *   Relatório Excel gerado em `relatorio_anexos.xlsx`.

***

## 📄 Exemplo de Nome de Arquivo

    2025.11.18 14.32 - Continente.pdf

***

## ✅ Próximos Passos

*   [ ] Adicionar opção para execução automática (sem GUI).
*   [ ] Envio de relatório por e-mail.
*   [ ] Configuração dinâmica via arquivo `.ini`.

***

## 📜 Licença

Este projeto é distribuído sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.