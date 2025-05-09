
# Sistema Automatizado de Cobrança – Python + Power BI + SAP

Este projeto demonstra a automação de cobrança baseada em dados extraídos do SAP, com integração a Python e visualização estratégica no Power BI.

## 💡 Objetivo
Automatizar o processo de cobrança reativa utilizando Python, com foco em:
- Identificação de clientes inadimplentes
- Envio automático de e-mails personalizados
- Visualização de métricas em Power BI
- Interface gráfica para execução de ações

## 🧩 Estrutura do Projeto

### 📁 SCRIPTS/
- `interface.py`: interface visual criada com PySide6
- `reativa.cobranca.py`: script para envio automatizado de cobranças reativas com base na tabela Aging

### 📁 INPUT/
- `AGING.png`: demonstração visual da base Aging usada como referência

### 📁 INTERFACE/
- `INTERFACE.png`: print da interface de execução do sistema

### 📁 DOCS/
- `DASHBOARD.png`: visão executiva do relatório Power BI
- `Guia para instalar Python.pdf`: instruções de instalação e uso

## ⚙️ Tecnologias Utilizadas
- Python 3.10.11
- Pandas
- PySide6
- smtplib + Outlook
- Power BI (relatórios)
- SAP (origem da base de dados)

## 📌 Observações
- Os scripts usam dados fictícios nas demonstrações visuais
- A versão do Python precisa ser 3.10.11 para compatibilidade
- Requer configuração prévia do Outlook para envios

---

© Projeto desenvolvido por Bruno Ricardo como demonstração técnica de automação de processos no ciclo de Order to Cash.
