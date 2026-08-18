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
    "saipos_pgto": "G",
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


SITE_LOJA_FOLDER = "Japa Nobre - Ouro preto"  # pasta raiz da loja dentro de "Shared Documents"


def _listar_filhos(token, site_id, caminho):
    """Lista os itens (arquivos/pastas) dentro de um caminho, usando o drive item por path."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{caminho}:/children"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json().get("value", [])


def localizar_arquivo_dashboard(token, site_id, mes_num, ano):
    """
    Localiza o arquivo do Dashboard do mês navegando pela estrutura de pastas
    (em vez de usar o endpoint /search, que retorna 403 com permissão de aplicativo).

    Estrutura esperada: Japa Nobre - Ouro preto/Financeiro/{ano}/{pasta do mês}/{arquivo}.xlsx
    A pasta do mês tem grafia inconsistente ("08- Agosto" vs "08 - Agosto"), então
    localizamos por correspondência parcial do nome do mês, não pelo nome exato.
    """
    mes_nome = MESES_PT[mes_num].split(" - ")[1]  # ex: "Agosto"

    caminho_ano = f"{SITE_LOJA_FOLDER}/Financeiro/{ano}"
    pastas_do_ano = _listar_filhos(token, site_id, caminho_ano)

    pasta_mes = None
    for item in pastas_do_ano:
        if "folder" in item and mes_nome.lower() in item["name"].lower():
            pasta_mes = item["name"]
            break

    if not pasta_mes:
        raise FileNotFoundError(
            f"Não encontrei a pasta do mês '{mes_nome}' dentro de '{caminho_ano}'. "
            f"Pastas disponíveis: {[i['name'] for i in pastas_do_ano]}"
        )

    caminho_mes = f"{caminho_ano}/{pasta_mes}"
    arquivos_do_mes = _listar_filhos(token, site_id, caminho_mes)

    for item in arquivos_do_mes:
        nome = item["name"].lower()
        if nome.endswith(".xlsx") and "dashboard" in nome and "ouro preto" in nome:
            return item["id"]

    raise FileNotFoundError(
        f"Não encontrei o arquivo Dashboard dentro de '{caminho_mes}'. "
        f"Arquivos disponíveis: {[i['name'] for i in arquivos_do_mes]}"
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


def determinar_linha_e_turno(token, site_id, item_id, linha_data):
    """Verifica se a linha da data já tem valores lançados.
    Se estiver vazia, é o 1º fechamento do dia (DIA), escreve nela mesma.
    Se já tiver algo, é o 2º fechamento do dia (NOITE), escreve na linha de baixo."""
    sheet = "Fechamento de Caixa"
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
        f"/workbook/worksheets('{sheet}')/range(address='B{linha_data}:J{linha_data}')"
    )
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    valores = resp.json()["values"][0]
    tem_valor = any(v not in (None, "", 0) for v in valores)

    if not tem_valor:
        return linha_data, "DIA"
    return linha_data + 1, "NOITE"


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

    # Observação (texto) — inclui a hora do fechamento junto
    hora = payload.get("hora_fechamento")
    obs_original = payload.get("obs_caixa", "")
    partes_obs = []
    if hora:
        partes_obs.append(f"Fechado às {hora}")
    if obs_original:
        partes_obs.append(obs_original)
    texto_obs = " — ".join(partes_obs)

    if texto_obs:
        endereco = f"{COLUNA_OBSERVACAO}{linha}"
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}"
            f"/workbook/worksheets('{sheet}')/range(address='{endereco}')"
        )
        resp = requests.patch(url, headers=headers, json={"values": [[texto_obs]]})
        resp.raise_for_status()

    # Nota: Total (K) e Diferença (L) já são calculados por fórmula na
    # própria planilha (=SUM(B:J) etc.) — não precisamos escrever ali.


# ─────────────────────────────────────────────
# ENVIAR E-MAIL PARA O JOÃO
# ─────────────────────────────────────────────
def _montar_anexo(nome_arquivo, data_uri):
    """Converte uma data URI ('data:image/jpeg;base64,...') em um anexo
    no formato que a Graph API espera para sendMail."""
    if not data_uri or "," not in data_uri:
        return None
    cabecalho, base64_puro = data_uri.split(",", 1)
    content_type = "image/jpeg"
    if "image/png" in cabecalho:
        content_type = "image/png"
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": nome_arquivo,
        "contentType": content_type,
        "contentBytes": base64_puro,
    }


def enviar_email_conferencia(token, payload):
    farol_emoji = {"verde": "🟢", "amarelo": "🟡", "vermelho": "🔴"}
    emoji = farol_emoji.get(payload["semaforo"], "⚪")

    # Extrai só o nome da cidade (depois do travessão), ex: "Japa Nobre — Ouro Preto" -> "Ouro Preto"
    cidade = payload["loja"].split("—")[-1].strip()
    turno = payload.get("turno", "")
    turno_label = f" {turno}" if turno else ""

    ano_f, mes_f, dia_f = payload["data"].split("-")
    data_formatada = f"{dia_f}/{mes_f}/{ano_f}"

    corpo_html = f"""
    <h2>{emoji} Fechamento de Caixa — {payload['loja']}{turno_label}</h2>
    <p><b>Data:</b> {payload['data']}</p>
    <p><b>Hora do fechamento:</b> {payload.get('hora_fechamento', '—')}</p>
    <p><b>Responsável:</b> {payload['responsavel']}</p>
    <p><b>Total contado:</b> R$ {payload['total_contado']:.2f}</p>
    <p><b>Total Saipos (conferência):</b> R$ {payload['total_saipos']:.2f}</p>
    <p><b>Diferença:</b> R$ {payload['diferenca']:.2f} — {emoji}</p>
    {"<p><b>Observação:</b> " + payload["obs_caixa"] + "</p>" if payload.get("obs_caixa") else ""}
    <p><i>Fotos do fechamento em anexo (papel do Saipos e máquinas de cartão).</i></p>
    """

    anexos = []
    anexo_saipos = _montar_anexo(
        f"saipos-{payload['data']}.jpg", payload.get("foto_saipos_base64")
    )
    anexo_maquina = _montar_anexo(
        f"maquinas-{payload['data']}.jpg", payload.get("foto_maquina_base64")
    )
    if anexo_saipos:
        anexos.append(anexo_saipos)
    if anexo_maquina:
        anexos.append(anexo_maquina)

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_REMETENTE_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "message": {
            "subject": f"{emoji} Fechamento Caixa {cidade}{turno_label} - {data_formatada}",
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": EMAIL_DESTINATARIO}}],
            "attachments": anexos,
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
    linha_data = encontrar_linha_da_data(token, site_id, item_id, data_fechamento)
    linha, turno = determinar_linha_e_turno(token, site_id, item_id, linha_data)
    payload["turno"] = turno
    preencher_linha(token, site_id, item_id, linha, payload)

    # Envio de e-mail imediato DESATIVADO de propósito — agora só o
    # job_conferencia_8h.py manda e-mail, 1x por dia, já consolidando
    # Dia + Noite + Saipos + PagSeguro num único e-mail. Isso evita
    # duplicar e-mails (um na hora do fechamento, outro na conciliação).

    return {"status": "ok", "linha_preenchida": linha, "turno": turno, "arquivo_id": item_id}
