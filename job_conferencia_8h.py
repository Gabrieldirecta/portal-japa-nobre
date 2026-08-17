"""
Job de Conciliação Diária — PagSeguro EDI vs Fechamento de Caixa
===================================================================

Roda 1x por dia via GitHub Actions (ver .github/workflows/conferencia-diaria.yml).

O que este script faz:
1. Calcula a data de ontem (D-1) — é o dia que o PagSeguro já tem
   disponível e é também o dia cujo fechamento de caixa já foi feito.
2. Busca no PagSeguro EDI o total de vendas em cartão (débito + crédito)
   daquele dia, usando o endpoint "transactional".
3. Busca na planilha "Fechamento de Caixa" (SharePoint) o que o gerente
   declarou nas colunas Débito e Crédito para o mesmo dia.
4. Compara os dois valores e classifica com farol (verde/amarelo/vermelho).
5. Envia um e-mail de resumo para o time financeiro.

Piloto: só a loja Ouro Preto por enquanto. Para expandir para outras
lojas, replicar as credenciais (USER/TOKEN do PagSeguro + pasta do
SharePoint) por loja — ver seção CONFIGURAÇÃO abaixo.

Variáveis de ambiente esperadas (configuradas como GitHub Secrets):
- AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
- PAGSEGURO_USER_OUROPRETO, PAGSEGURO_TOKEN_OUROPRETO
- EMAIL_REMETENTE_MAILBOX (mailbox Microsoft 365 usada para enviar)
- EMAIL_DESTINATARIO (para quem vai o resumo)
"""

import os
from datetime import date, timedelta

import requests
from requests.auth import HTTPBasicAuth

# ─────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

SITE_HOSTNAME = "hakunabt.sharepoint.com"
SITE_PATH = "/sites/Teste-JapaNobre-Ouropreto"
SITE_LOJA_FOLDER = "Japa Nobre - Ouro preto"

EMAIL_REMETENTE_MAILBOX = os.environ.get("EMAIL_REMETENTE_MAILBOX", "gabriel@hakunabatata.com.br")
EMAIL_DESTINATARIO = os.environ.get("EMAIL_DESTINATARIO", "gestaojapanobre@gmail.com")

PAGSEGURO_USER = os.environ["PAGSEGURO_USER_OUROPRETO"]
PAGSEGURO_TOKEN = os.environ["PAGSEGURO_TOKEN_OUROPRETO"]

MESES_PT = {
    1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Marco", 4: "04 - Abril",
    5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
    9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro",
}


# ─────────────────────────────────────────────
# PAGSEGURO EDI
# ─────────────────────────────────────────────
def buscar_total_cartao_pagseguro(data_str):
    """Busca o total de vendas em cartão (débito + crédito) no PagSeguro EDI
    para a data informada (formato AAAA-MM-DD). Retorna (total, validado)."""
    url = f"https://edi.api.pagbank.com.br/movement/v3.00/transactional/{data_str}"
    params = {"pageNumber": 1, "pageSize": 1000}
    resp = requests.get(url, params=params, auth=HTTPBasicAuth(PAGSEGURO_USER, PAGSEGURO_TOKEN))
    resp.raise_for_status()

    validado = resp.headers.get("VALIDADO", "false").lower() == "true"
    dados = resp.json()
    detalhes = dados.get("detalhes", [])

    total = sum(float(item.get("valor_total_transacao", 0)) for item in detalhes)
    return total, validado


# ─────────────────────────────────────────────
# MICROSOFT GRAPH (SharePoint / Excel)
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


def obter_site_id(token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_HOSTNAME}:{SITE_PATH}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]


def _listar_filhos(token, site_id, caminho):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{caminho}:/children"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json().get("value", [])


def localizar_arquivo_dashboard(token, site_id, mes_num, ano):
    mes_nome = MESES_PT[mes_num].split(" - ")[1]
    caminho_ano = f"{SITE_LOJA_FOLDER}/Financeiro/{ano}"
    pastas_do_ano = _listar_filhos(token, site_id, caminho_ano)
    pasta_mes = None
    for item in pastas_do_ano:
        if "folder" in item and mes_nome.lower() in item["name"].lower():
            pasta_mes = item["name"]
            break
    if not pasta_mes:
        raise FileNotFoundError(f"Nao encontrei a pasta do mes '{mes_nome}'")
    caminho_mes = f"{caminho_ano}/{pasta_mes}"
    arquivos_do_mes = _listar_filhos(token, site_id, caminho_mes)
    for item in arquivos_do_mes:
        nome = item["name"].lower()
        if nome.endswith(".xlsx") and "dashboard" in nome and "ouro preto" in nome:
            return item["id"]
    raise FileNotFoundError(f"Nao encontrei o arquivo Dashboard dentro de '{caminho_mes}'")


from datetime import datetime as _dt


def _celula_para_data(valor):
    try:
        return _dt.strptime(str(valor), "%m/%d/%Y").date()
    except ValueError:
        pass
    try:
        epoch = _dt(1899, 12, 30)
        return (epoch + timedelta(days=float(valor))).date()
    except (ValueError, TypeError):
        return None


def buscar_debito_credito_planilha(token, site_id, item_id, data_alvo):
    """Lê as colunas A (data), D (débito) e E (crédito) e retorna a soma
    débito+crédito da linha correspondente à data."""
    sheet = "Fechamento de Caixa"
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/worksheets('{sheet}')/range(address='A1:E61')"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    valores = resp.json()["values"]

    for linha in valores:
        celula_data = linha[0] if len(linha) > 0 else None
        if not celula_data:
            continue
        data_linha = _celula_para_data(celula_data)
        if data_linha == data_alvo:
            debito = float(linha[3]) if len(linha) > 3 and linha[3] else 0.0
            credito = float(linha[4]) if len(linha) > 4 and linha[4] else 0.0
            return debito + credito

    return None  # não encontrou a linha (dia ainda não fechado)


# ─────────────────────────────────────────────
# E-MAIL DE RESUMO
# ─────────────────────────────────────────────
def enviar_email_resumo(token, data_str, total_pagseguro, total_planilha, validado):
    if total_planilha is None:
        farol_emoji = "⚪"
        mensagem_diff = "Fechamento de caixa daquele dia ainda não foi encontrado na planilha."
        diferenca_txt = "—"
    else:
        diferenca = total_pagseguro - total_planilha
        a = abs(diferenca)
        if a <= 5:
            farol_emoji, mensagem_diff = "🟢", "Bateu certinho"
        elif a <= 30:
            farol_emoji, mensagem_diff = "🟡", "Pequena diferença — verificar"
        else:
            farol_emoji, mensagem_diff = "🔴", "Diferença relevante — verificar com urgência"
        diferenca_txt = f"R$ {diferenca:.2f}"

    aviso_validado = "" if validado else "<p style='color:#c00'><b>⚠️ Atenção:</b> o PagSeguro ainda não confirmou que os dados desse dia estão totalmente processados (VALIDADO=false). Os números abaixo podem estar incompletos.</p>"

    corpo_html = f"""
    <h2>{farol_emoji} Conciliação PagSeguro x Fechamento — Ouro Preto</h2>
    <p><b>Data conferida:</b> {data_str}</p>
    {aviso_validado}
    <p><b>Total em cartão (PagSeguro EDI):</b> R$ {total_pagseguro:.2f}</p>
    <p><b>Total Débito+Crédito (declarado no fechamento):</b> {"R$ %.2f" % total_planilha if total_planilha is not None else "não encontrado"}</p>
    <p><b>Diferença:</b> {diferenca_txt} — {mensagem_diff}</p>
    """

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_REMETENTE_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "message": {
            "subject": f"{farol_emoji} Conciliação PagSeguro — Ouro Preto — {data_str}",
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": EMAIL_DESTINATARIO}}],
        }
    }
    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()


# ─────────────────────────────────────────────
# EXECUÇÃO PRINCIPAL
# ─────────────────────────────────────────────
def main():
    ontem = date.today() - timedelta(days=1)
    data_str = ontem.strftime("%Y-%m-%d")
    print(f"Conciliando o dia: {data_str}")

    total_pagseguro, validado = buscar_total_cartao_pagseguro(data_str)
    print(f"Total PagSeguro (cartão): R$ {total_pagseguro:.2f} — validado={validado}")

    token = obter_token_graph()
    site_id = obter_site_id(token)
    item_id = localizar_arquivo_dashboard(token, site_id, ontem.month, ontem.year)
    total_planilha = buscar_debito_credito_planilha(token, site_id, item_id, ontem)
    print(f"Total planilha (Débito+Crédito): {total_planilha}")

    enviar_email_resumo(token, data_str, total_pagseguro, total_planilha, validado)
    print("E-mail de resumo enviado.")


if __name__ == "__main__":
    main()
