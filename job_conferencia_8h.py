"""
Job de Conciliação Diária — PagSeguro EDI vs Fechamento de Caixa (por turno)
==============================================================================

Roda 1x por dia via GitHub Actions (ver .github/workflows/conferencia-diaria.yml).

O que este script faz:
1. Calcula a data de ontem (D-1).
2. Lê na planilha "Fechamento de Caixa" as linhas daquele dia — pode ser
   1 linha (loja sem 2 turnos) ou 2 linhas (turno DIA + turno NOITE),
   pegando a hora exata de cada fechamento (gravada na Observação pelo
   formulário, ex: "Fechado às 15:30").
3. Busca as transações do PagSeguro EDI daquele dia (com horário exato
   de cada uma) e separa quais caem dentro da janela de cada turno:
     - DIA: da meia-noite até a hora do fechamento do turno DIA
     - NOITE: da hora do fechamento do turno DIA até a hora do
       fechamento do turno NOITE — mesmo que isso atravesse a meia-noite
       (nesse caso, busca também as transações do dia seguinte)
4. Compara o total de cada turno com o Débito+Crédito declarado no
   formulário daquele turno, e envia um e-mail de resumo por turno.

Piloto: só a loja Ouro Preto por enquanto.

Variáveis de ambiente esperadas (GitHub Secrets):
- AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
- PAGSEGURO_USER_OUROPRETO, PAGSEGURO_TOKEN_OUROPRETO
- EMAIL_REMETENTE_MAILBOX, EMAIL_DESTINATARIO
"""

import os
import re
from datetime import date, datetime, timedelta

import requests
from requests.auth import HTTPBasicAuth

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

CIDADE = "Ouro Preto"

MESES_PT = {
    1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Marco", 4: "04 - Abril",
    5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
    9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro",
}


def parse_minutos(hora_str):
    partes = hora_str.split(":")
    return int(partes[0]) * 60 + int(partes[1])


def _extrair_hora_fechamento(obs_texto):
    if not obs_texto:
        return None
    m = re.search(r"Fechado às (\d{2}:\d{2})", obs_texto)
    return m.group(1) if m else None


def buscar_transacoes_pagseguro(data_str):
    url = f"https://edi.api.pagbank.com.br/movement/v3.00/transactional/{data_str}"
    params = {"pageNumber": 1, "pageSize": 1000}
    resp = requests.get(url, params=params, auth=HTTPBasicAuth(PAGSEGURO_USER, PAGSEGURO_TOKEN))
    resp.raise_for_status()

    validado = resp.headers.get("VALIDADO", "false").lower() == "true"
    detalhes = resp.json().get("detalhes", [])
    transacoes = [
        (item.get("hora_inicial_transacao", "00:00:00"), float(item.get("valor_total_transacao", 0)))
        for item in detalhes
    ]
    return transacoes, validado


def somar_janela(transacoes, minuto_inicio, minuto_fim):
    total = 0.0
    for hora, valor in transacoes:
        m = parse_minutos(hora)
        if minuto_inicio <= m < minuto_fim:
            total += valor
    return total


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


def _celula_para_data(valor):
    try:
        return datetime.strptime(str(valor), "%m/%d/%Y").date()
    except ValueError:
        pass
    try:
        epoch = datetime(1899, 12, 30)
        return (epoch + timedelta(days=float(valor))).date()
    except (ValueError, TypeError):
        return None


def buscar_linhas_do_dia(token, site_id, item_id, data_alvo):
    sheet = "Fechamento de Caixa"
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/worksheets('{sheet}')/range(address='A1:M61')"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    valores = resp.json()["values"]

    linhas_encontradas = []
    for i, linha in enumerate(valores):
        celula_data = linha[0] if len(linha) > 0 else None
        if not celula_data:
            continue
        data_linha = _celula_para_data(celula_data)
        if data_linha == data_alvo:
            linhas_encontradas.append(linha)
            if i + 1 < len(valores):
                proxima = valores[i + 1]
                tem_data_propria = len(proxima) > 0 and proxima[0]
                if not tem_data_propria:
                    linhas_encontradas.append(proxima)
            break

    resultado = []
    for linha in linhas_encontradas:
        debito = float(linha[3]) if len(linha) > 3 and linha[3] else 0.0
        credito = float(linha[4]) if len(linha) > 4 and linha[4] else 0.0
        obs = linha[12] if len(linha) > 12 else ""
        resultado.append({
            "debito_credito": debito + credito,
            "hora_fechamento": _extrair_hora_fechamento(obs),
        })

    if len(resultado) == 1:
        resultado[0]["turno"] = None
    elif len(resultado) >= 2:
        resultado[0]["turno"] = "DIA"
        resultado[1]["turno"] = "NOITE"

    return resultado


def enviar_email_resumo(token, data_str, turno, total_pagseguro, total_planilha, validado):
    ano_f, mes_f, dia_f = data_str.split("-")
    data_formatada = f"{dia_f}/{mes_f}/{ano_f}"
    turno_label = f" {turno}" if turno else ""

    if total_pagseguro is None:
        farol_emoji = "⚪"
        corpo_extra = "<p>Não foi possível conciliar automaticamente este turno — hora de fechamento não encontrada na planilha.</p>"
        diferenca_txt = "—"
        total_pagseguro_txt = "não calculado"
    elif total_planilha is None:
        farol_emoji = "⚪"
        corpo_extra = "<p>Fechamento de caixa daquele dia/turno ainda não foi encontrado na planilha.</p>"
        diferenca_txt = "—"
        total_pagseguro_txt = f"R$ {total_pagseguro:.2f}"
    else:
        diferenca = total_pagseguro - total_planilha
        a = abs(diferenca)
        if a <= 5:
            farol_emoji, msg = "🟢", "Bateu certinho"
        elif a <= 30:
            farol_emoji, msg = "🟡", "Pequena diferença — verificar"
        else:
            farol_emoji, msg = "🔴", "Diferença relevante — verificar com urgência"
        diferenca_txt = f"R$ {diferenca:.2f} — {msg}"
        total_pagseguro_txt = f"R$ {total_pagseguro:.2f}"
        corpo_extra = ""

    aviso_validado = "" if validado else "<p style='color:#c00'><b>⚠️ Atenção:</b> o PagSeguro ainda não confirmou que os dados desse período estão totalmente processados (VALIDADO=false). Os números abaixo podem estar incompletos.</p>"

    corpo_html = f"""
    <h2>{farol_emoji} Fechamento Caixa {CIDADE}{turno_label} - {data_formatada}</h2>
    {aviso_validado}
    {corpo_extra}
    <p><b>Total em cartão (PagSeguro EDI):</b> {total_pagseguro_txt}</p>
    <p><b>Total Débito+Crédito (declarado no fechamento):</b> {"R$ %.2f" % total_planilha if total_planilha is not None else "não encontrado"}</p>
    <p><b>Diferença:</b> {diferenca_txt}</p>
    """

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_REMETENTE_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "message": {
            "subject": f"{farol_emoji} Fechamento Caixa {CIDADE}{turno_label} - {data_formatada}",
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": EMAIL_DESTINATARIO}}],
        }
    }
    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()


def main():
    ontem = date.today() - timedelta(days=1)
    data_str = ontem.strftime("%Y-%m-%d")
    print(f"Conciliando o dia: {data_str}")

    token = obter_token_graph()
    site_id = obter_site_id(token)
    item_id = localizar_arquivo_dashboard(token, site_id, ontem.month, ontem.year)
    linhas = buscar_linhas_do_dia(token, site_id, item_id, ontem)
    print(f"Turnos encontrados na planilha: {len(linhas)}")

    transacoes_hoje, validado_hoje = buscar_transacoes_pagseguro(data_str)
    print(f"Transações PagSeguro em {data_str}: {len(transacoes_hoje)} — validado={validado_hoje}")

    if not linhas:
        print("Nenhuma linha encontrada para esse dia na planilha. Encerrando.")
        return

    if len(linhas) == 1:
        total_pagseguro = sum(v for _, v in transacoes_hoje)
        enviar_email_resumo(token, data_str, None, total_pagseguro, linhas[0]["debito_credito"], validado_hoje)
        print("E-mail enviado (sem separação de turno).")
        return

    dia_info, noite_info = linhas[0], linhas[1]
    hora_dia = dia_info["hora_fechamento"]
    hora_noite = noite_info["hora_fechamento"]

    if hora_dia:
        m_dia = parse_minutos(hora_dia)
        total_dia = somar_janela(transacoes_hoje, 0, m_dia)
        enviar_email_resumo(token, data_str, "DIA", total_dia, dia_info["debito_credito"], validado_hoje)
        print(f"Turno DIA: PagSeguro R$ {total_dia:.2f} | Planilha R$ {dia_info['debito_credito']:.2f}")
    else:
        enviar_email_resumo(token, data_str, "DIA", None, dia_info["debito_credito"], validado_hoje)
        print("Turno DIA: sem hora de fechamento registrada, não foi possível conciliar.")

    if hora_dia and hora_noite:
        m_dia = parse_minutos(hora_dia)
        m_noite = parse_minutos(hora_noite)
        cruzou_meia_noite = m_noite <= m_dia

        total_noite = somar_janela(transacoes_hoje, m_dia, 24 * 60)
        validado_noite = validado_hoje

        if cruzou_meia_noite:
            amanha = ontem + timedelta(days=1)
            amanha_str = amanha.strftime("%Y-%m-%d")
            print(f"Turno NOITE cruza a meia-noite — buscando também {amanha_str}")
            transacoes_amanha, validado_amanha = buscar_transacoes_pagseguro(amanha_str)
            total_noite += somar_janela(transacoes_amanha, 0, m_noite)
            validado_noite = validado_hoje and validado_amanha

        enviar_email_resumo(token, data_str, "NOITE", total_noite, noite_info["debito_credito"], validado_noite)
        print(f"Turno NOITE: PagSeguro R$ {total_noite:.2f} | Planilha R$ {noite_info['debito_credito']:.2f}")
    else:
        enviar_email_resumo(token, data_str, "NOITE", None, noite_info["debito_credito"], validado_hoje)
        print("Turno NOITE: sem hora de fechamento registrada, não foi possível conciliar.")

    print("Concluído.")


if __name__ == "__main__":
    main()
