# 📦 Marketplace Order SLA Tracker & Automated Reporter

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458.svg)](https://pandas.pydata.org/)
[![SendGrid](https://img.shields.io/badge/SendGrid-API-blue.svg)](https://sendgrid.com/)

---

## 🇧🇷 Português

### 📄 Descrição

Pipeline automatizado desenvolvido para monitoramento contínuo de SLA de pedidos pendentes em plataformas de Marketplace. O script processa dados operacionais, calcula prazos em dias úteis, identifica itens em risco ou em atraso e dispara relatórios executivos em HTML via SendGrid API, com planilhas Excel dinâmicas anexadas em memória.

### 🚀 Funcionalidades

* **Cálculo de SLA em Dias Úteis:** Utilização da biblioteca `pandas.tseries.offsets.BDay` para monitorar prazos operacionais.
* **Relatório HTML Responsivo:** Geração de e-mails formatados com indicadores de status (`🟢 No prazo`, `🟡 Risco`, `🔴 Atrasado`) e tabelas responsivas.
* **Anexo Dinâmico em Memória (`io.BytesIO`):** Criação de planilhas Excel com múltiplas abas diretamente na memória, sem necessidade de gravação em disco.
* **Notificação Automatizada:** Integração com a SendGrid API para envio de relatórios à equipe de logística e operações.

### 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **Pandas & OpenPyXL** — análise, manipulação e geração de relatórios
* **SendGrid API** — envio automatizado de e-mails
* **PyTZ** — tratamento e conversão de fusos horários

---

## 🇺🇸 English

### 📄 Description

Automated data pipeline built for continuous tracking of pending marketplace order SLAs. The script processes operational data, calculates business-day deadlines, identifies at-risk or delayed items, and dispatches executive HTML email reports via the SendGrid API with dynamically generated in-memory Excel attachments.

### 🚀 Key Features

* **Business Day SLA Engine:** Uses `pandas.tseries.offsets.BDay` to calculate business-day deadlines and operational thresholds.
* **Responsive HTML Email Reporting:** Generates formatted HTML emails with KPI status indicators (`🟢 On Time`, `🟡 At Risk`, `🔴 Delayed`) and responsive tables.
* **In-Memory Excel Attachment (`io.BytesIO`):** Builds multi-sheet Excel workbooks directly in memory without writing temporary files to disk.
* **Automated Notification:** Integrates with the SendGrid API for targeted delivery to logistics and operations teams.

### 🛠️ Tech Stack

* **Python 3.x**
* **Pandas & OpenPyXL** — data processing and reporting
* **SendGrid API** — transactional email delivery
* **PyTZ** — timezone handling and conversion

---

## ⚙️ Configuração e Execução / Setup & Execution

### 1. Clonar o repositório / Clone the repository

```bash
git clone https://github.com/IsabelaKalid/mktp-sla-reporter_automation.git
cd mktp-sla-reporter_automation
```

### 2. Criar ambiente virtual / Create a virtual environment

```bash
python -m venv venv
```

Windows:

```bash
venv\Scripts\activate
```

### 3. Instalar dependências / Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configurar variáveis de ambiente / Configure environment variables

Crie um arquivo `.env` na raiz do projeto e configure as credenciais necessárias para a integração com a SendGrid API.

**Nunca compartilhe credenciais reais no GitHub.**

### 5. Executar / Run

```bash
python mktp-sla-reporter_automation.py
```

---

## 🔐 Segurança / Security

Credenciais, chaves de API e outras informações sensíveis não devem ser armazenadas diretamente no código ou versionadas no Git.

Recomenda-se utilizar variáveis de ambiente para armazenar informações sensíveis.

---

## 🎯 Objetivo / Objective

### 🇧🇷 Português

Automatizar o acompanhamento de SLA de pedidos e transformar dados operacionais em informações acionáveis, reduzindo o trabalho manual e permitindo que as equipes de logística e operações identifiquem rapidamente pedidos em risco ou atrasados.

### 🇺🇸 English

Automate marketplace order SLA monitoring and transform operational data into actionable information, reducing manual work and enabling logistics and operations teams to quickly identify at-risk or delayed orders.

---

## ⚠️ Disclaimer

Este projeto foi desenvolvido para fins de demonstração de automação, processamento de dados, geração de relatórios e integração com APIs.

This project was developed for demonstration purposes, focusing on automation, data processing, reporting, and API integration.
