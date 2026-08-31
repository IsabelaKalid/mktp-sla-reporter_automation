import base64
import io
import os
from datetime import datetime, timedelta
import numpy as np
import openpyxl
import pandas as pd
from pandas.tseries.offsets import BDay  # Business Day = dia útil / Business Day calculation
import pytz
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Attachment,
    Disposition,
    FileContent,
    FileName,
    FileType,
    From,
    Mail,
)

# =====================================================
# CONFIGURAÇÕES E PARÂMETROS / CONFIGURATION & PARAMETERS
# =====================================================

# SharePoint parameters (Generic variables for security)
# Parâmetros do SharePoint (Variáveis genéricas por segurança)
HOSTNAME = os.getenv("SHAREPOINT_HOST", "empresa.sharepoint.com")
SITE_PATH = os.getenv("SHAREPOINT_SITE", "/sites/empresa.internacional")
DOCUMENT_LIBRARY = "Documentos"
FILE_PATH = "CONTROLE_MKTP/PEDIDOS_SELLER_CENTER_V2.xlsx"
SHEET_NAME = "DESCRICAO COMPRA"

# List of email recipients (Generic placeholders)
# Lista de e-mails de destino (Substitutos genéricos)
EMAIL_DESTINO = [
    "gestao_operacoes@empresa.com.br",
    "logistica@empresa.com.br",
    "supervisao@empresa.com.br"
]

# Mapping table columns to email headers
# Mapeamento das colunas da tabela de origem para as colunas do e-mail
COLUNAS_TABELA = [
    "Data da compra",
    "SKU Produto",
    "Descrição do Material",
    "Ordem de Cliente",
    "Ordem de Compra",
    "Sequencial",
    "Quantidade",
    "Invoice",
    "Data Emissão NF (Chegada no CD)",
    "Data Entrega Cliente"
]

COLUNAS_EMAIL = [
    "data_compra",
    "sku_produto",
    "descricao_material",
    "ordem_cliente",
    "ordem_compra",
    "sequencial",
    "quantidade",
    "invoice",
    "Chegada_CD",
    "Entrega_Limite" 
]

COLUNAS_FATURADOS_D1 = [
    "data_compra",
    "sku_produto",
    "descricao_material",
    "ordem_cliente",
    "ordem_compra",
    "sequencial",
    "quantidade",
    "invoice",
    "Data Entrega Cliente"
]

# =====================================================
# LEITURA E TRATAMENTO DE DADOS / DATA CLEANING & PREP
# =====================================================

# Load data from SharePoint (Simulated/Generic call)
# Carga dos dados do SharePoint (Chamada simulada/genérica)
# df_final, ok = read_excel_from_sharepoint(
#     hostname=HOSTNAME, site_path=SITE_PATH,
#     document_library=DOCUMENT_LIBRARY, file_path=FILE_PATH, sheet_name=SHEET_NAME
# )

# Extracting relevant sub-dataframe and setting headers
# Extraindo o sub-dataframe relevante e redefinindo os cabeçalhos
df = df_final.iloc[6:, 1:52].copy()
df.columns = df.iloc[0]
df = df.iloc[1:].reset_index(drop=True)

# Convert all object columns to string to prevent PyArrow type mismatch errors
# Converte todas as colunas do tipo objeto para string evitando conflitos de tipo
for col in df.select_dtypes(include='object').columns:
    df[col] = df[col].astype(str)

colunas_data = [
    "Data da compra",
    "Data Emissão NF (Chegada no CD)",
    "Data Entrega Cliente"
]
 
for col in colunas_data:
    df[col] = df[col].replace("00:00:00", np.nan)
    df[col] = pd.to_datetime(df[col], errors="coerce")

hoje = pd.Timestamp.today().normalize()

# Filter open orders that are purchased but pending delivery
# Filtra pedidos abertos/comprados sem data de entrega registrada
df_abertos = df[
    (df["Sts de Compra"].str.upper() == "COMPRADO") &
    (df["Data Entrega Cliente"].isnull())
].copy()

# Month-Year grouping logic / Lógica de agrupamento por Mês e Ano
df_abertos["Mes_Ref"] = df_abertos["Data da compra"].dt.to_period("M").dt.to_timestamp()
df_abertos = df_abertos.sort_values(by=["Mes_Ref", "Data da compra"], ascending=[True, True])
 
meses_pt = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}
df_abertos["Mes_Ano"] = df_abertos["Mes_Ref"].dt.month.map(meses_pt) + "/" + df_abertos["Mes_Ref"].dt.year.astype(str)
renomear_email_dict = dict(zip(COLUNAS_TABELA[:-1], COLUNAS_EMAIL))
df_abertos_renomeado = df_abertos.rename(columns=renomear_email_dict)

# =====================================================
# CÁLCULO DE PREDICADOS E SLA / SLA & DELAY COMPUTATION
# =====================================================

hoje_date = pd.Timestamp.today().normalize()

def calcular_status(row):
    """
    Computes visual indicator for SLA state based on business days.
    Calcula o indicador visual do SLA baseado em dias úteis decorridos.
    """
    dias_uteis = len(pd.bdate_range(start=row["data_compra"], end=hoje_date))
    if dias_uteis <= 20:
        return "🟢"
    elif 21 <= dias_uteis <= 25:
        return "🟡"
    else:
        return "🔴"
 
df_abertos_status = df_abertos_renomeado.copy()
df_abertos_status["Prazo"] = df_abertos_status.apply(calcular_status, axis=1)

# Set deadline limit to 25 business days / Define limite de entrega em 25 dias úteis
df_abertos_status["Entrega_Limite"] = df_abertos_status["data_compra"] + BDay(25)

# Calculate days delayed / Calcula quantidade de dias em atraso
df_abertos_status["Dias em Atraso"] = (hoje_date - df_abertos_status["Entrega_Limite"]).dt.days
df_abertos_status["Dias em Atraso"] = df_abertos_status["Dias em Atraso"].apply(lambda x: x if x > 0 else 0)

# Flag delayed status / Marca flag de atraso para formatação
df_abertos_status["Atrasado"] = (df_abertos_status["Prazo"] == "🔴") 

# Reorder columns: SLA status right next to purchase date
# Reorganiza colunas: insere o status de Prazo logo após a data de compra
cols = df_abertos_status.columns.tolist()
idx = cols.index("data_compra")
cols.insert(idx + 1, cols.pop(cols.index("Prazo")))
df_abertos_status = df_abertos_status[cols]

def formatar_datas_para_email(df_input):
    """
    Formats standard pandas datetimes to DD/MM/YYYY.
    Formata objetos datetime do pandas para o formato padrão DD/MM/YYYY.
    """
    df_copy = df_input.copy()
    for col in df_copy.columns:
        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            df_copy[col] = df_copy[col].dt.strftime("%d/%m/%Y")
    return df_copy

COLUNAS_EMAIL_COM_STATUS = ["Prazo"] + COLUNAS_EMAIL + ["Atrasado", "Dias em Atraso"]

# Grouping pending items by Month-Year for email rendering
# Agrupa os pedidos pendentes por Mês/Ano para montar as seções do e-mail
abertos_por_mes = {
    mes: formatar_datas_para_email(
        dados[COLUNAS_EMAIL_COM_STATUS].reset_index(drop=True)
    )
    for mes, dados in df_abertos_status.groupby("Mes_Ano", sort=False)
}

# =====================================================
# PEDIDOS FATURADOS D-1 / DELIVERED ORDERS D-1
# =====================================================

fuso_manaus = pytz.timezone("America/Manaus")
agora_manaus = datetime.now(fuso_manaus)
hoje_local = agora_manaus.date()
ontem_local = hoje_local - timedelta(days=1)

df["Data_Entrega_Data"] = df["Data Entrega Cliente"].dt.normalize().dt.date

COLUNAS_FATURADOS_ORIGEM = [
    "Data da compra",
    "SKU Produto",
    "Descrição do Material",
    "Ordem de Cliente",
    "Ordem de Compra",
    "Quantidade",
    "Data Entrega Cliente"
]

df_faturados_d1 = df.loc[
    df["Data_Entrega_Data"] == ontem_local,
    COLUNAS_FATURADOS_ORIGEM
].copy()

df_faturados_d1 = df_faturados_d1.rename(columns={
    "Data da compra": "Data_Compra",
    "SKU Produto": "Sku",
    "Descrição do Material": "Descrição",
    "Ordem de Cliente": "Ordem_Cliente",
    "Ordem de Compra": "Ordem_Compra",
    "Quantidade": "Quantidade",
    "Data Entrega Cliente": "Data_Entrega"
})

df_faturados_d1 = formatar_datas_para_email(df_faturados_d1)

# Summary of orders delivered in the last 30 days
# Resumo de pedidos entregues/concluídos nos últimos 30 dias
df_concluidos_30dias = df[
    (df["Sts de Compra"].str.upper() == "ENTREGUE") &
    (df["Data Entrega Cliente"].dt.date >= hoje_local - timedelta(days=30))
]
qtd_concluidos_30dias = len(df_concluidos_30dias)

resumo_status_email = {
    '🟢No prazo (até 20 dias úteis)': df_abertos_status["Prazo"].value_counts().get('🟢', 0),
    '🟡Risco de atraso (21 a 25 dias úteis)': df_abertos_status["Prazo"].value_counts().get('🟡', 0),
    '🔴Atrasado (mais de 25 dias úteis)': df_abertos_status["Prazo"].value_counts().get('🔴', 0),
} 

# =====================================================
# MONTAGEM DO LAYOUT HTML / HTML TEMPLATE BUILDER
# =====================================================

def dataframe_para_html(df_input):
    """
    Renders styled pandas DataFrame to HTML table with conditional highlights.
    Converte o DataFrame em tabela HTML estilizada com destaques condicionais.
    """
    df_html = df_input.reset_index(drop=True).copy()
 
    # Highlight deadline if delayed / Destaca a data limite se houver atraso
    if "Entrega_Limite" in df_html.columns and "Atrasado" in df_html.columns:
        df_html["Entrega_Limite"] = df_html.apply(
            lambda row: (
                f"<span style='color:#d32f2f; font-weight:bold;'>{row['Entrega_Limite']}</span>"
                if row["Atrasado"] else row["Entrega_Limite"]
            ),
            axis=1
        )
 
    # Highlight delay counter / Destaca contador de dias em atraso
    if "Dias em Atraso" in df_html.columns:
        df_html["Dias em Atraso"] = df_html["Dias em Atraso"].apply(
            lambda x: (
                f"<span style='color:#8e0000; font-weight:bold;'>{x}</span>"
                if pd.notna(x) and int(x) > 0 else "0"
            )
        )
 
    if "Atrasado" in df_html.columns:
        df_html = df_html.drop(columns=["Atrasado"])
 
    return df_html.to_html(
        index=False, border=0, justify="center", classes="tabela-pedidos", escape=False
    )

# Building executive HTML layout / Montagem da estrutura visual do e-mail
html_content = f"""
<html>
<head>
<style>
.div-tabela {{ overflow-x: auto; margin-bottom: 20px; }}
table.tabela-pedidos {{ border-collapse: collapse; font-size: 10px; width: 100%; }}
table.tabela-pedidos th {{ background-color: #f0f4fa; padding: 6px; border: 1px solid #cccccc; text-align: center; font-size: 9px; }}
table.tabela-pedidos td {{ padding: 8px; border: 1px solid #cccccc; vertical-align: top; font-size: 9px; }}
</style>
</head>
<body style="margin:0; padding:0; font-family: Arial, sans-serif; background-color:#f4f6f8;">
<table width="100%" cellpadding="0" cellspacing="0">
<tr>
<td align="center">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:1200px; margin:0 auto; background:#ffffff; border-radius:12px; overflow:hidden;">
<tr>
<td style="background:#0066cc; padding:20px; text-align:center;">
    <h2 style="color:#ffffff; margin:0;">📦 Relatório de Pedidos Pendentes de Faturamento - <b>{hoje_local.strftime('%d/%m/%Y')}</b></h2>
</td>
</tr>
<tr>
<td style="padding:20px; color:#333333; font-size:12px;">
    <p>Prezados,</p> 
    <p>Este e-mail apresenta o status atualizado dos pedidos pendentes de faturamento, organizados por mês de compra e SLA, bem como os pedidos faturados no dia anterior (D-1).</p>
</td>
</tr>
<tr>
<td style="padding:0 20px 20px 20px;">
    <table width="100%" cellpadding="10" cellspacing="0" style="border:1px solid #dddddd; border-radius:10px;">
        <tr style="background:#f0f4fa; text-align:center; font-size:12px;">
            <td><b>🟢 No prazo (até 20 dias úteis)</b><br>{resumo_status_email['🟢No prazo (até 20 dias úteis)']}</td>
            <td><b>🟡 Risco (21 a 25 dias úteis)</b><br>{resumo_status_email['🟡Risco de atraso (21 a 25 dias úteis)']}</td>
            <td><b>🔴 Atrasado (+ de 25 dias úteis)</b><br>{resumo_status_email['🔴Atrasado (mais de 25 dias úteis)']}</td>
            <td><b>🚚 Concluídos (últimos 30 dias)</b><br>{qtd_concluidos_30dias}</td>
        </tr>
    </table>
</td>
</tr>
"""

# Append tables per month / Anexa as tabelas agrupadas por mês
for mes, dados in abertos_por_mes.items():
    qtd_pedidos = len(dados)
    dados_clean = dados.copy()
    dados_clean.columns.name = None
    html_content += f"""
    <tr>
    <td style="padding:20px;">
        <table width="100%" cellpadding="10" cellspacing="0" style="border:1px solid #cccccc; border-radius:10px;">
            <tr>
                <td style="background:#eef3fb; padding:5px 10px;">
                    <b>{mes}</b> - <span style="color:red; font-weight:bold;">[{qtd_pedidos} pedidos]</span>
                </td>
            </tr>
            <tr>
                <td><div class="div-tabela">{dataframe_para_html(dados_clean)}</div></td>
            </tr>
        </table>
    </td>
    </tr>
    """

# Append D-1 delivered orders table / Anexa a tabela de faturados ontem (D-1)
html_content += f"""
<tr>
<td style="padding:20px;">
    <table width="100%" cellpadding="10" cellspacing="0" style="border:1px solid #cccccc; border-radius:10px;">
        <tr>
            <td style="background:#eef3fb; border-radius:10px 10px 0 0;">
                <b>📦 Pedidos faturados ontem - ({ontem_local.strftime('%d/%m/%Y')})</b>
            </td>
        </tr>
        <tr>
            <td>
                <div class="div-tabela">
"""

if df_faturados_d1.empty:
    html_content += "<p style='font-size:11px;'>Nenhum pedido faturado no dia anterior.</p>"
else:
    COLUNAS_EMAIL_RESUMO = [
        "Data_Compra", "Sku", "Descrição", "Ordem_Cliente",
        "Ordem_Compra", "Quantidade", "Data_Entrega"
    ]
    df_faturados_clean = df_faturados_d1[COLUNAS_EMAIL_RESUMO].copy()
    df_faturados_clean.columns.name = None
    html_content += dataframe_para_html(df_faturados_clean)

html_content += """
                </div>
            </td>
        </tr>
    </table>
</td>
</tr>
<tr>
<td style="padding:15px; text-align:center; font-size:11px; color:#777;">
    Relatório gerado automaticamente via pipeline de dados • Automação de Processos
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
</html>
"""

# =====================================================
# ENVIO VIA SENDGRID API / SENDGRID DISPATCH
# =====================================================

def nome_aba_valido(nome):
    """Sanitizes sheet names for Excel compatibility / Sanitiza nomes de abas do Excel."""
    return nome.replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "").replace("[", "").replace("]", "").replace(":", "-")

if not abertos_por_mes or all(dados.empty for dados in abertos_por_mes.values()):
    print("⚠️ Nenhum pedido pendente encontrado para gerar o relatório.")
else:
    # Build Excel File dynamically in RAM memory / Gera arquivo Excel dinamicamente na RAM
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        for mes, dados in abertos_por_mes.items():
            aba_valida = nome_aba_valido(mes)
            dados_excel = dados.drop(columns=["Atrasado"], errors="ignore")
            dados_excel.to_excel(writer, sheet_name=aba_valida, index=False)

    excel_buffer.seek(0)
    encoded_excel = base64.b64encode(excel_buffer.read()).decode('utf-8')

    # Create email attachment / Cria o anexo para o e-mail
    anexo_excel = Attachment(
        FileContent(encoded_excel),
        FileName(f"Pedidos_Pendentes_[{hoje_local.strftime('%d-%m-%Y')}].xlsx"),
        FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        Disposition("attachment"),
    )

    sendgrid_api_key = os.getenv("SENDGRID_API_KEY")
    sender_email = os.getenv("SENDER_EMAIL", "automacao@empresa.com.br")

    message = Mail(
        from_email=From(sender_email, 'Automação de Relatórios'),
        to_emails=EMAIL_DESTINO,
        subject=f"🚚 Relatório Marketplace – Pedidos Pendentes - {hoje_local.strftime('%d/%m/%Y')}",
        html_content=html_content
    )
    message.add_attachment(anexo_excel)

    # Dispatch via SendGrid API / Dispara mensagem via API do SendGrid
    if sendgrid_api_key:
        sg = SendGridAPIClient(sendgrid_api_key)
        response = sg.send(message)
        print(f"✅ E-mail enviado com sucesso! Status Code: {response.status_code}")
    else:
        print("⚠️ Chave SENDGRID_API_KEY não localizada no ambiente. Verifique suas variáveis.")