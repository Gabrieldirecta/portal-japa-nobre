"""
Módulo de conferência de fechamento de caixa — Japa Nobre
===========================================================

O que este script faz, a partir do JSON que já chega no webhook
(receber-fechamento):

1. Calcula o mês/ano da data do fechamento e monta o nome do
   Dashboard esperado (ex: "09 - Setembro Dashboard Ouro Preto 2026.xlsx").
2. Localiza esse arquivo no SharePoint via Microsoft Graph (busca por
   nome, não por caminho fixo — assim não quebra se a pasta tiver
   variação de espaço/hífen).
3. Abre a aba "Fechamento de Caixa", encontra a LINHA cuja data bate
   com a do formulário, e preenche só as células dessa linha (sem
   mexer nas fórmulas existentes de soma/semana).
4. Envia um e-mail para gestaojapanobre@gmail.com (João) com o
   resumo e o farol (verde/amarelo/vermelho).

⚠️ IMPORTANTE — antes de rodar de verdade:
- Este script usa autenticação "client credentials" (App Registration
  já existente: portal-japa-nobre-app, permissões Files.ReadWrite.All
  e Mail.Send — ambas já concedidas/"Grant admin consent" feito).
- O MAPEAMENTO de colunas abaixo foi deduzido lendo a estrutura da
  planilha de Setembro. No primeiro teste real, precisamos conferir
  se cada valor caiu na coluna certa.
- Este ambiente (chat) não tem rede liberada para graph.microsoft.com,
  então este script não foi executado aqui — ele deve rodar dentro da
  Azure Function, que já tem essa permissão de rede.
"""

import os
import requests
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────
TENANT_ID = "3b1e307a-e1ce-465f-bf8a-92d18b94dd57"
CLIENT_ID = "1f079fc7-b408-472e-a753-d8270381428d"
CLIENT_SECRET = os.environ["PORTAL_APP_CLIENT_SECRET"]  # guardado como app setting no Azure

SITE_HOSTNAME = "hakunabt.sharepoint.com"
SITE_PATH = "/sites/Teste-JapaNobre-Ouropreto"

EMAIL_DESTINATARIO = os.environ.get("EMAIL_CONFERENCIA", "gestaojapanobre@gmail.com")
EMAIL_REMETENTE_MAILBOX = "gabriel@hakunabatata.com.br"  # mailbox usada para enviar (precisa licença + permissão Mail.Send)

MESES_PT = {
    1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março", 4: "04 - Abril",
    5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
    9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro",
}

# Mapeamento: chave do payload (meios_pagamento) -> coluna da tabela Fechamento
# Ordem real das colunas na planilha (B até J):
# B=Dinheiro Sistema, C=Cortesia, D=Débito, E=Crédito,
# F=Pgt. Online(ifood), G=Saipos, H=Vale Refeição, I=Voucher, J=PIX
COLUNA_MEIO = {
    "dinheiro": "B",
    "cortesia": "C",
    "debito": "D",
    "credito": "E",
    # "online" não vem do formulário simplificado (é preenchido pelo escritório)
    "saipos": "G",
    "vale": "H",
    "voucher": "I",
    "pix": "J",
}
COLUNA_TOTAL = "K"
COLUNA_DIFERENCA = "L"
COLUNA_OBSERVACAO = "M"


# ─────────────────────────────────────────────
# AUTENTICAÇÃO
# ─────────────────────────────────────────────
def obter_token_graph():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data)
    resp.raise_for_status()
    return resp.json()["access_token"]


# ─────────────────────────────────────────────
# LOCALIZAR O SITE E O ARQUIVO DO MÊS
# ─────────────────────────────────────────────
def obter_site_id(token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_PATH}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]


def localizar_arquivo_dashboard(token, site_id, mes_num, ano):
    """Busca o arquivo do Dashboard do mês pelo nome, sem depender do caminho exato da pasta."""
    mes_label = MESES_PT[mes_num]  # ex: "09 - Setembro"
    nome_esperado = f"{mes_label} Dashboard Ouro Preto {ano}.xlsx"

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{mes_label.split(' - ')[1]} Dashboard Ouro Preto {ano}')"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    resultados = resp.json().get("value", [])

    for item in resultados:
        if item["name"].lower() == nome_esperado.lower():
            return item["id"]

    raise FileNotFoundError(
        f"Não encontrei o arquivo '{nome_esperado}' no SharePoint. "
        f"Confirme se o Dashboard desse mês já foi criado/duplicado."
    )


# ─────────────────────────────────────────────
# ENCONTRAR A LINHA DA DATA NA TABELA
# ─────────────────────────────────────────────
def encontrar_linha_da_data(token, site_id, item_id, data_fechamento):
    """Lê a coluna A da aba 'Fechamento de Caixa' e retorna o número da linha que bate com a data."""
    sheet = "Fechamento de Caixa"
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
        f"/workbook/worksheets('{sheet}')/range(address='A1:A61')"
    )
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    valores = resp.json()["values"]

    data_alvo = datetime.strptime(data_fechamento, "%Y-%m-%d").date()

    for i, linha in enumerate(valores, start=1):
        celula = linha[0]
        if not celula:
            continue
        try:
            # Excel pode devolver como texto "9/1/2026" ou serial — tratamos texto primeiro
            data_celula = datetime.strptime(str(celula), "%m/%d/%Y").date()
        except ValueError:
            continue
        if data_celula == data_alvo:
            return i  # número da linha na planilha (1-indexed, igual ao Excel)

    raise ValueError(f"Não encontrei a linha correspondente a {data_fechamento} na aba '{sheet}'.")


# ─────────────────────────────────────────────
# ESCREVER OS VALORES NA LINHA ENCONTRADA
# ─────────────────────────────────────────────
def preencher_linha(token, site_id, item_id, linha, payload):
    sheet = "Fechamento de Caixa"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    meios = payload["meios_pagamento"]

    for chave, coluna in COLUNA_MEIO.items():
        valor = meios.get(chave, 0)
        if valor:  # só escreve se houver valor (preserva fórmulas/formatos em branco)
            endereco = f"{coluna}{linha}"
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
                f"/workbook/worksheets('{sheet}')/range(address='{endereco}')"
            )
            body = {"values": [[valor]]}
            resp = requests.patch(url, headers=headers, json=body)
            resp.raise_for_status()

    # Observação (texto)
    if payload.get("obs_caixa"):
        endereco = f"{COLUNA_OBSERVACAO}{linha}"
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
            f"/workbook/worksheets('{sheet}')/range(address='{endereco}')"
        )
        resp = requests.patch(url, headers=headers, json={"values": [[payload["obs_caixa"]]]})
        resp.raise_for_status()

    # Nota: Total (K) e Diferença (L) já são calculados por fórmula na
    # própria planilha (=SUM(B:J) etc.) — não precisamos escrever ali.


# ─────────────────────────────────────────────
# ENVIAR E-MAIL PARA O JOÃO
# ─────────────────────────────────────────────
def enviar_email_conferencia(token, payload):
    farol_emoji = {"verde": "🟢", "amarelo": "🟡", "vermelho": "🔴"}
    emoji = farol_emoji.get(payload["semaforo"], "⚪")

    corpo_html = f"""
    <h2>{emoji} Fechamento de Caixa — {payload['loja']}</h2>
    <p><b>Data:</b> {payload['data']}</p>
    <p><b>Responsável:</b> {payload['responsavel']}</p>
    <p><b>Total contado:</b> R$ {payload['total_contado']:.2f}</p>
    <p><b>Total Saipos (conferência):</b> R$ {payload['total_saipos']:.2f}</p>
    <p><b>Diferença:</b> R$ {payload['diferenca']:.2f} — {emoji}</p>
    {"<p><b>Observação:</b> " + payload["obs_caixa"] + "</p>" if payload.get("obs_caixa") else ""}
    """

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_REMETENTE_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "message": {
            "subject": f"{emoji} Fechamento {payload['loja']} — {payload['data']}",
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": EMAIL_DESTINATARIO}}],
        }
    }
    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()


# ─────────────────────────────────────────────
# FUNÇÃO PRINCIPAL — chamar isso a partir do receber-fechamento
# ─────────────────────────────────────────────
def processar_fechamento(payload: dict):
    data_fechamento = payload["data"]  # formato "AAAA-MM-DD"
    ano, mes, _ = data_fechamento.split("-")
    mes_num = int(mes)
    ano_num = int(ano)

    token = obter_token_graph()
    site_id = obter_site_id(token)
    item_id = localizar_arquivo_dashboard(token, site_id, mes_num, ano_num)
    linha = encontrar_linha_da_data(token, site_id, item_id, data_fechamento)
    preencher_linha(token, site_id, item_id, linha, payload)

    # Envio de e-mail: só funciona depois que a permissão Mail.Send
    # for concedida (Grant admin consent) no App Registration.
    # Se ainda não foi liberada, comente a linha abaixo para não
    # quebrar o restante do fluxo.
    enviar_email_conferencia(token, payload)

    return {"status": "ok", "linha_preenchida": linha, "arquivo_id": item_id}
