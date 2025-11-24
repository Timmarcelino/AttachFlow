# AttachFlow

AttachFlow é uma aplicação pessoal para automatizar a leitura de caixas de email via IMAP, identificar mensagens que obedecem a regras configuráveis e descarregar anexos para pastas específicas. O objectivo é simplificar rotinas manuais, criar históricos consistentes e fornecer um painel web de controlo, estatísticas e execução de regras.

Este repositório contém duas camadas principais:

- **attachflow_core** — Módulo independente com toda a lógica funcional (IMAP, regras, download, logging).
- **attachflow_web** — Interface web construída em Django para gestão de contas, regras, execuções, logs e estatísticas.

O desenvolvimento é pensado para evoluir de forma incremental, permitindo um MVP simples e funcional, seguido de uma expansão modular.

---

## 🧱 Estado Actual

- Repositório inicial criado.
- Planeamento arquitectural definido.
- README introduzido (documento vivo em actualização contínua).

---

## 🎯 Objectivos do Projecto

1. **Ler caixas de email via IMAP (Outlook/Microsoft, Gmail, etc.)**
2. **Aplicar regras configuráveis**:
   - Remetente
   - Assunto
   - Padrões de nome de ficheiro (regex)
   - Tipo de anexo
3. **Descarregar anexos automaticamente** para pastas definidas.
4. **Registar execuções e métricas**:
   - Nº de emails processados
   - Nº de anexos descarregados
   - Logs detalhados
5. **Oferecer um painel web em Django**:
   - CRUD de contas de email
   - CRUD de regras
   - Execução manual de regras
   - Histórico de execuções
   - Dashboard com estatísticas

---

## 🏗️ Estrutura Inicial (prevista)

```text
attachflow/
  attachflow_core/
    (lógica IMAP, regras, downloads)
  attachflow_web/
    (projecto Django)
  config/
  data/
  logs/
  README.md
  .venv/
  requirements.txt
