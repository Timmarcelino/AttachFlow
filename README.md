# AttachFlow

AttachFlow é uma aplicação pessoal para automatizar a leitura de caixas de email via IMAP, identificar mensagens segundo regras configuráveis e descarregar anexos para pastas específicas. Inclui um painel web baseado em Django para gestão das contas, regras, execuções, logs e estatísticas. O objetivo é transformar uma tarefa repetitiva num fluxo de captura inteligente, organizado e extensível.

---

## 🧩 Visão Geral

O projecto divide-se em dois módulos principais:

- **attachflow_core**  
  Núcleo independente responsável por:
  - Ligação IMAP
  - Teste de acesso e listagem de pastas
  - Execução das regras (pesquisa, filtragem, download de anexos)
  - Geração de nomes de ficheiros segundo templates
  - Estatísticas e resultados estruturados para persistência no Django

- **attachflow_web**  
  Painel de controlo em Django para:
  - Gerir contas de email
  - Configurar regras de captura
  - Guardar execuções, anexos processados e logs
  - Integrar com o núcleo `attachflow_core`
  - Evoluir para dashboards e scheduler

---

## 🏗️ Estrutura do Projeto

AttachFlow/
attachflow_core/
email_core.py
init.py

attachflow_web/
manage.py
attachflow_web/
settings.py
urls.py
...
accounts/
models.py
admin.py
...
rules/
models.py
admin.py
...

config/
.gitkeep
data/
.gitkeep
logs/
.gitkeep
docs/
.gitkeep
scripts/
.gitkeep

.venv/
requirements.txt
.gitignore
README.md


---

## ⚙️ Funcionalidades Implementadas

### 🔹 Núcleo (attachflow_core)

- Função de **teste de ligação IMAP** com devolução de sucesso/erro e lista de pastas.
- Função para **listar pastas IMAP**.
- Implementação completa de:
  - Pesquisa de emails via critérios (FROM / SUBJECT).
  - Filtragem por regex do nome do anexo.
  - Processamento seguro de anexos.
  - Geração de nomes de ficheiros com template configurável:
    ```
    {date:%Y.%m.%d %H.%M} - {rule_name}{index}{ext}
    ```
  - Evita conflitos entre ficheiros.
  - Marca emails como lidos e move-os para uma pasta final, se configurado.
- Estatísticas estruturadas (`RunStats`) para integração com Django.

---

### 🔹 Django – Apps já criadas

#### **accounts**
Modelo `EmailAccount`:
- Configuração de contas IMAP
- Campos: host, port, username, password, SSL, pasta base, activo
- Registo no Admin

#### **rules**
Modelo `Rule`:
- Relacionada com `EmailAccount`
- Filtros por remetente, assunto e regex do anexo
- Pasta IMAP específica por regra
- Pasta local de destino
- Template de nome de ficheiro
- Flags: mark_as_read e move_to_folder
- Registo no Admin

---

## 🔧 Funcionalidades Planeadas e Aprovadas

### Para `EmailAccount`
- Guardar lista de pastas IMAP (`cached_folders`)
- Historizar estado da ligação:
  - `last_connection_ok`
  - `last_connection_message`

### Para `Rule`
- Campo final `filename_template`
- Suporte a placeholders:
  - `{date:%FMT}`, `{rule_name}`, `{account_name}`,
    `{original_name}`, `{index}`, `{ext}`

### Modelos adicionais
- **ProcessedEmail**  
  Para evitar processar o mesmo UID novamente.
- **JobExecution**  
  Para históricos de execuções.
- **AttachmentLog** (com estatísticas detalhadas).

### Integração Django ↔ Core
- Management Command:

python manage.py run_rule <id_da_regra>
- Versão multi-regra:
python manage.py run_all_rules


### Painel Web (fases posteriores)
- Botão “Testar ligação”
- Selector de pastas IMAP após teste
- Navegação de pastas (dropdown / árvore)
- Execução manual de regras
- Dashboard com estatísticas (Chart.js)
- Scheduler simples (cron / APScheduler)

---

## 🧪 Ambiente de Desenvolvimento

```bash
# Criar ambiente virtual
python -m venv .venv

# Activar
.venv\Scripts\activate   # Windows
source .venv/bin/activate  # Linux/macOS

# Instalar dependências
pip install -r requirements.txt

📦 Dependências Principais
   Python 3.12+
   Django 5.1+
   IMAPClient
   PyYAML
   APScheduler
   regex
   requests

🚀 Roadmap Atual

   ☑   Estrutura base do projeto
   ☑   Setup do Django + apps iniciais
   ☑   Implementação do núcleo IMAP completo
   ☐   Adicionar campos extras às models (cached_folders, etc.)
   ☐   Criar modelo ProcessedEmail
   ☐   Criar modelo JobExecution & AttachmentLog
   ☐   Implementar ligação Django → Core (run_rule)
   ☐   Botão de “Testar Ligação”
   ☐   Selector de pastas IMAP
   ☐   Dashboard inicial
   ☐   Scheduler interno (opcional)
   ☐   UI refinada

📄 Licença
   Uso pessoal e experimental.

👤 Autor
   Criado e mantido por Valentim M. Pinto.