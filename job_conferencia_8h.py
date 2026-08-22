"""
Job de Conciliacao Diaria - Ouro Preto
=========================================

Roda 1x por dia via GitHub Actions.

Fontes cruzadas:
- Saipos (o "Sistema") - fonte principal de comparacao para cada forma de
  pagamento, incluindo o relatorio detalhado de vendas em Dinheiro.
- Planilha (o que o gerente contou fisicamente e lancou no formulario).
- PagSeguro - confere de forma independente Debito, Credito, PIX e
  Voucher/Vale Refeicao (que realmente passam pela maquininha).
- Banco Inter - usado so para detectar excecoes (PIX que caiu direto no
  CNPJ, fora do fluxo normal do PagSeguro).

Categorizacao dos pagamentos do Saipos (baseado em testes reais):
- "Dinheiro"                  -> Dinheiro (secao propria, com relatorio)
- "Debito"/"Credito"          -> Cartao (cruza com PagSeguro)
- "Pix"                       -> PIX (cruza com PagSeguro)
- Vale Refeicao real (Sodexo, VR, Alelo, Ticket, "Vale Refeicao")
                               -> Voucher/Vale (cruza com PagSeguro)
- "Pago Online" e "Voucher Parceiro Desconto" (repasse iFood/99Food)
                               -> Online/Parceiros (so Saipos x Planilha,
                                  nao passa por nenhum banco/adquirente)
- "Cortesia"                  -> Cortesia (so Saipos x Planilha)
- Qualquer outro nome         -> Outros (aparece no log para revisao)

ATENCAO - premissas ainda nao 100% confirmadas com dados reais:
- O nome exato que o PagSeguro usa para Voucher/Vale Refeicao no campo
  'arranjo_ur' ainda nao foi visto em um teste real (so vimos DEBIT_*,
  CREDIT_* e PIX ate agora). Usamos uma lista de palpites (ver eh_voucher).
  Se aparecer 'DESCONHECIDO' na coluna PagSeguro do Voucher, precisamos
  ajustar isso com um teste real assim que houver uma venda em voucher.
"""

import os
import re
import time
from datetime import date, datetime, timedelta

import requests
from requests.auth import HTTPBasicAuth

# ─────────────────────────────────────────────
# CONFIGURACAO
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

INTER_CLIENT_ID = os.environ["INTER_CLIENT_ID_OUROPRETO"]
INTER_CLIENT_SECRET = os.environ["INTER_CLIENT_SECRET_OUROPRETO"]
INTER_CERT_PATH = os.environ.get("INTER_CERT_PATH", "inter_cert.pem")
INTER_KEY_PATH = os.environ.get("INTER_KEY_PATH", "inter_key.pem")

SAIPOS_TOKEN = os.environ["SAIPOS_TOKEN"]
print(f"[DEBUG] SAIPOS_TOKEN: tamanho={len(SAIPOS_TOKEN)} inicio='{SAIPOS_TOKEN[:15]}' fim='{SAIPOS_TOKEN[-10:]}'")
SAIPOS_ID_STORE_OUROPRETO = 71180

CIDADE = "Ouro Preto"

MESES_PT = {
    1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Marco", 4: "04 - Abril",
    5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
    9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro",
}

# Colunas da planilha "Fechamento de Caixa" (A=data, depois os meios)
COLUNA_PLANILHA = {
    "dinheiro": "B",
    "cortesia": "C",
    "debito": "D",
    "credito": "E",
    "online": "F",
    "saipos_proprio": "G",
    "vale": "H",
    "voucher": "I",
    "pix": "J",
}


# ─────────────────────────────────────────────
# HELPERS DE HORA
# ─────────────────────────────────────────────
def parse_minutos(hora_str):
    partes = hora_str.split(":")
    return int(partes[0]) * 60 + int(partes[1])


def _extrair_hora_fechamento(obs_texto):
    if not obs_texto:
        return None
    # Tolerante a "Fechado às" (com acento) e "Fechado as" (sem acento) -
    # a Azure Function grava sem acento, mas mantemos os dois por segurança.
    m = re.search(r"Fechado \w*s? (\d{2}:\d{2})", obs_texto)
    return m.group(1) if m else None


# ─────────────────────────────────────────────
# PAGSEGURO EDI
# ─────────────────────────────────────────────
def buscar_transacoes_pagseguro(data_str):
    url = f"https://edi.api.pagbank.com.br/movement/v3.00/transactional/{data_str}"
    params = {"pageNumber": 1, "pageSize": 1000}
    resp = requests.get(url, params=params, auth=HTTPBasicAuth(PAGSEGURO_USER, PAGSEGURO_TOKEN))
    resp.raise_for_status()

    validado = resp.headers.get("VALIDADO", "false").lower() == "true"
    detalhes = resp.json().get("detalhes", [])
    transacoes = [
        (
            item.get("hora_inicial_transacao", "00:00:00"),
            float(item.get("valor_total_transacao", 0)),
            item.get("arranjo_ur", "DESCONHECIDO"),
        )
        for item in detalhes
    ]
    return transacoes, validado


def eh_debito(arranjo):
    return arranjo.startswith("DEBIT_")


def eh_credito(arranjo):
    return arranjo.startswith("CREDIT_")


def eh_pix(arranjo):
    return arranjo == "PIX"


def eh_voucher(arranjo):
    """Palpite ainda nao confirmado com dados reais - ajustar quando
    tivermos um dia com venda em voucher/vale refeicao no PagSeguro."""
    a = arranjo.upper()
    return any(p in a for p in ["VOUCHER", "VR", "VA", "ALELO", "SODEXO", "TICKET"])


def somar_janela(transacoes, minuto_inicio, minuto_fim, filtro=None):
    total = 0.0
    for hora, valor, arranjo in transacoes:
        m = parse_minutos(hora)
        if minuto_inicio <= m < minuto_fim:
            if filtro is None or filtro(arranjo):
                total += valor
    return total


# ─────────────────────────────────────────────
# BANCO INTER (so para excecoes de PIX)
# ─────────────────────────────────────────────
def obter_token_inter():
    url = "https://cdpj.partners.bancointer.com.br/oauth/v2/token"
    resp = requests.post(
        url,
        data={
            "client_id": INTER_CLIENT_ID,
            "client_secret": INTER_CLIENT_SECRET,
            "grant_type": "client_credentials",
            "scope": "extrato.read",
        },
        cert=(INTER_CERT_PATH, INTER_KEY_PATH),
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def buscar_pix_recebido_inter(data_str):
    token = obter_token_inter()
    url = "https://cdpj.partners.bancointer.com.br/banking/v2/extrato"
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params={"dataInicio": data_str, "dataFim": data_str},
        cert=(INTER_CERT_PATH, INTER_KEY_PATH),
    )
    resp.raise_for_status()
    transacoes = resp.json().get("transacoes", [])

    total_pix = 0.0
    for t in transacoes:
        if t.get("tipoTransacao", "").upper() == "PIX" and t.get("tipoOperacao", "").upper() == "C":
            total_pix += float(t.get("valor", 0))
    return total_pix


def checar_excecao_pix_inter(data_str, total_pix_pagseguro):
    try:
        total_pix_inter = buscar_pix_recebido_inter(data_str)
    except Exception as e:
        print(f"Nao foi possivel checar excecoes no Banco Inter: {e}")
        return None

    diferenca = total_pix_inter - total_pix_pagseguro
    print(f"Checagem de excecao - Inter: R$ {total_pix_inter:.2f} | PagSeguro: R$ {total_pix_pagseguro:.2f}")

    if diferenca > 5:
        print(f"Possivel PIX direto no CNPJ detectado: R$ {diferenca:.2f}")
        return (total_pix_inter, total_pix_pagseguro, diferenca)
    return None


# ─────────────────────────────────────────────
# SAIPOS
# ─────────────────────────────────────────────
def categorizar_pagamento_saipos(desc):
    """Classifica o texto livre do Saipos (desc_store_payment_type) em
    uma das categorias que usamos na conciliacao."""
    d = (desc or "").lower()
    if "dinheiro" in d:
        return "dinheiro"
    if "cortesia" in d:
        return "cortesia"
    if "crédito" in d or "credito" in d:
        return "credito"
    if "débito" in d or "debito" in d:
        return "debito"
    if "pix" in d:
        return "pix"
    if "voucher parceiro" in d or "pago online" in d or "online" in d:
        return "online_parceiro"
    if "vale" in d or "voucher" in d or "sodexo" in d or "alelo" in d or "ticket" in d:
        return "voucher_vale"
    return "outros"


def buscar_vendas_saipos(data_str):
    """Busca todas as vendas do dia (todas as lojas), com paginacao e
    nova tentativa em caso de erro (403/504) da API do Saipos.

    Parametros de retry/pausa alinhados com outro projeto que já
    funciona de forma estável com essa mesma API: 5 tentativas com
    espera crescente (20s, 40s, 60s, 80s, 100s), timeout de 120s por
    chamada, e pausa de 5s entre páginas de paginação — isso evita
    estourar o limite de consultas por minuto do lado do Saipos."""
    url = "https://data.saipos.io/v1/search_sales"
    headers = {"Authorization": f"Bearer {SAIPOS_TOKEN}"}

    vendas = []
    offset = 0
    while True:
        params = {
            "p_date_column_filter": "shift_date",
            "p_filter_date_start": f"{data_str}T00:00:00",
            "p_filter_date_end": f"{data_str}T23:59:59",
            "p_limit": 1000,
            "p_offset": offset,
        }
        for tentativa in range(5):
            resp = requests.get(url, headers=headers, params=params, timeout=120)
            if resp.status_code == 200:
                break
            espera = 20 * (tentativa + 1)  # 20s, 40s, 60s, 80s, 100s
            print(f"Saipos respondeu {resp.status_code}, tentativa {tentativa + 1}/5 — aguardando {espera}s: {resp.text[:200]}")
            time.sleep(espera)
        else:
            raise RuntimeError(f"Saipos nao respondeu 200 apos 5 tentativas (offset={offset})")

        time.sleep(5)  # pausa entre páginas, para nao disparar limite de consultas

        pagina = resp.json()
        if not isinstance(pagina, list):
            raise RuntimeError(f"Resposta inesperada do Saipos: {resp.text[:300]}")

        vendas.extend(pagina)
        if len(pagina) < 1000:
            break
        offset += 1000

    return [v for v in vendas if v.get("id_store") == SAIPOS_ID_STORE_OUROPRETO and v.get("canceled") != "S"]


def montar_dados_saipos_por_turno(data_str):
    """Retorna um dict por turno ('Dia'/'Noite'/'Desconhecido'), cada um com:
    - categorias: {categoria: valor_total}
    - dinheiro_detalhe: lista de {pedido, hora, valor} das vendas em dinheiro
    - cortesia_detalhe: lista de {pedido, hora, valor} das vendas em cortesia
    """
    vendas = buscar_vendas_saipos(data_str)

    dados = {}
    for v in vendas:
        turno = (v.get("store_shift") or {}).get("desc_store_shift", "Desconhecido")
        if turno not in dados:
            dados[turno] = {"categorias": {}, "dinheiro_detalhe": [], "cortesia_detalhe": []}

        for p in (v.get("payments") or []):
            desc = p.get("desc_store_payment_type", "")
            valor = float(p.get("payment_amount", 0))
            categoria = categorizar_pagamento_saipos(desc)

            dados[turno]["categorias"][categoria] = dados[turno]["categorias"].get(categoria, 0.0) + valor

            item_detalhe = {
                "pedido": v.get("sale_number") or v.get("id_sale"),
                "hora": (p.get("created_at") or v.get("created_at") or "")[11:16],
                "valor": valor,
            }

            if categoria == "dinheiro":
                dados[turno]["dinheiro_detalhe"].append(item_detalhe)
            elif categoria == "cortesia":
                dados[turno]["cortesia_detalhe"].append(item_detalhe)

            if categoria == "outros":
                print(f"Forma de pagamento nao categorizada: '{desc}' (venda {v.get('sale_number')})")

    return dados


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
    """Le a aba Fechamento de Caixa (colunas A ate M) e retorna 1 ou 2
    dicts (1 por turno), cada um com os valores de cada coluna da planilha."""
    sheet = "Fechamento de Caixa"
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/worksheets('{sheet}')/range(address='A1:M75')"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    valores = resp.json()["values"]

    linhas_encontradas = []
    linhas_numeros = []
    for i, linha in enumerate(valores):
        celula_data = linha[0] if len(linha) > 0 else None
        if not celula_data:
            continue
        data_linha = _celula_para_data(celula_data)
        if data_linha == data_alvo:
            linhas_encontradas.append(linha)
            linhas_numeros.append(i + 1)  # linha real na planilha (1-indexed)
            if i + 1 < len(valores):
                proxima = valores[i + 1]
                tem_data_propria = len(proxima) > 0 and proxima[0]
                if not tem_data_propria:
                    linhas_encontradas.append(proxima)
                    linhas_numeros.append(i + 2)
            break

    def valor_coluna(linha, letra):
        indices = {"B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9}
        idx = indices[letra]
        return float(linha[idx]) if len(linha) > idx and linha[idx] else 0.0

    resultado = []
    for linha, num_linha in zip(linhas_encontradas, linhas_numeros):
        planilha = {chave: valor_coluna(linha, col) for chave, col in COLUNA_PLANILHA.items()}
        total_contado = float(linha[10]) if len(linha) > 10 and linha[10] else None
        obs = linha[12] if len(linha) > 12 else ""
        resultado.append({
            "linha": num_linha,
            "planilha": planilha,
            "total_contado": total_contado,
            "hora_fechamento": _extrair_hora_fechamento(obs),
            "observacao": obs,
        })

    if len(resultado) == 1:
        resultado[0]["turno"] = None
    elif len(resultado) >= 2:
        resultado[0]["turno"] = "DIA"
        resultado[1]["turno"] = "NOITE"

    return resultado


def escrever_valor_online_planilha(token, site_id, item_id, linha, valor_saipos):
    """Escreve automaticamente o total do Saipos na coluna F (Online/Parceiros)
    da linha do turno, fechando o preenchimento que antes dependia do
    escritório digitar manualmente."""
    sheet = "Fechamento de Caixa"
    endereco = f"F{linha}"
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/workbook/worksheets('{sheet}')/range(address='{endereco}')"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.patch(url, headers=headers, json={"values": [[valor_saipos]]})
    resp.raise_for_status()


# ─────────────────────────────────────────────
# MONTAGEM DO E-MAIL
# ─────────────────────────────────────────────
def _classificar_farol(diferenca, limite_verde=5, limite_amarelo=30):
    a = abs(diferenca)
    if a <= limite_verde:
        return "verde", "OK"
    elif a <= limite_amarelo:
        return "amarelo", "Verificar"
    else:
        return "vermelho", "Urgente"


CORES_FAROL = {
    "verde": ("#E8F8EF", "#27AE60"),
    "amarelo": ("#FDF3E7", "#E67E22"),
    "vermelho": ("#FBE9E7", "#E74C3C"),
    "cinza": ("#F0F0F0", "#999999"),
}


def _linha_tabela(fonte, saipos_val, planilha_val, pagseguro_val=None, informativo=False):
    """Monta uma linha <tr> da tabela principal. pagseguro_val=None significa
    'nao aplicavel' (mostra traco). Se informativo=True, a linha nao entra no
    calculo do farol (ex: Online/Parceiros, que a loja sempre lanca como 0
    porque nao tem acesso a esse valor - quem preenche depois e o escritorio).
    Retorna (html, farol_ou_none)."""
    if informativo:
        cor_fundo, cor_texto = CORES_FAROL["cinza"]
        status = "Preenchido auto."
        diferenca_txt = "—"
        farol = None
    elif saipos_val is None or planilha_val is None:
        cor_fundo, cor_texto = CORES_FAROL["cinza"]
        status = "N/D"
        diferenca_txt = "—"
        farol = None
    else:
        diferenca = saipos_val - planilha_val
        farol, status = _classificar_farol(diferenca)
        cor_fundo, cor_texto = CORES_FAROL[farol]
        diferenca_txt = f"R$ {diferenca:.2f}"

    pagseguro_txt = f"R$ {pagseguro_val:.2f}" if pagseguro_val is not None else "—"
    saipos_txt = f"R$ {saipos_val:.2f}" if saipos_val is not None else "—"
    planilha_txt = f"R$ {planilha_val:.2f}" if planilha_val is not None else "—"

    html = f"""
    <tr>
      <td style="padding:8px 10px;border:1px solid #ddd;">{fonte}</td>
      <td align="right" style="padding:8px 10px;border:1px solid #ddd;">{planilha_txt}</td>
      <td align="right" style="padding:8px 10px;border:1px solid #ddd;">{saipos_txt}</td>
      <td align="right" style="padding:8px 10px;border:1px solid #ddd;color:#888;">{pagseguro_txt}</td>
      <td align="center" style="padding:8px 10px;border:1px solid #ddd;background:{cor_fundo};color:{cor_texto};font-weight:700;">{status}</td>
      <td align="right" style="padding:8px 10px;border:1px solid #ddd;">{diferenca_txt}</td>
    </tr>
    """
    return html, farol


def montar_secao_turno(turno_label, planilha, saipos_categorias, pagseguro_por_categoria):
    """Monta as linhas da tabela principal para um turno (exclui Dinheiro e
    Cortesia, que tem secao propria com relatorio detalhado).
    Retorna (html, lista_de_farois) - farois so das linhas que contam
    para o farol geral (Online/Parceiros fica de fora, por ser informativo)."""
    linhas_html = []
    farois = []

    mapeamentos = [
        ("Débito", "debito", "debito", "debito", False),
        ("Crédito", "credito", "credito", "credito", False),
        ("PIX", "pix", "pix", "pix", False),
        ("Voucher/Vale Refeição", "voucher_vale", "vale_voucher_soma", "voucher", False),
        ("Online/Parceiros (iFood, 99Food etc.)", "online_parceiro", "online", None, True),
    ]

    for fonte_label, cat_saipos, cat_planilha, cat_pagseguro, informativo in mapeamentos:
        saipos_val = saipos_categorias.get(cat_saipos, 0.0)

        if cat_planilha == "vale_voucher_soma":
            planilha_val = planilha.get("vale", 0.0) + planilha.get("voucher", 0.0)
        else:
            planilha_val = planilha.get(cat_planilha, 0.0)

        pagseguro_val = pagseguro_por_categoria.get(cat_pagseguro) if cat_pagseguro else None

        html, farol = _linha_tabela(fonte_label, saipos_val, planilha_val, pagseguro_val, informativo=informativo)
        linhas_html.append(html)
        if farol:
            farois.append(farol)

    nota_online = """
    <p style="font-size:11px;color:#999;margin:4px 0 0;">
      * Online/Parceiros: a loja não tem acesso a esse valor no fechamento (sempre lança 0).
      O valor mostrado na coluna Saipos acima já foi preenchido automaticamente na planilha
      Dashboard pelo sistema. Essa linha não entra no status/farol do fechamento.
    </p>
    """

    tabela = f"""
    <h3 style="margin:18px 0 6px;color:#333;">{"Turno " + turno_label if turno_label else "Fechamento do dia"}</h3>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
      <tr style="background:#F0F0F0;">
        <th align="left" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Fonte</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Loja</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Saipos</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">PagSeguro</th>
        <th align="center" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Status</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Diferença</th>
      </tr>
      {"".join(linhas_html)}
    </table>
    {nota_online}
    """
    return tabela, farois


def _montar_secao_com_detalhe(titulo, turno_label, planilha_val, saipos_val, detalhe, mostrar_deposito=False):
    """Funcao generica usada tanto para Dinheiro quanto para Cortesia -
    mostra o total (Loja x Saipos) e a lista detalhada de pedidos/hora.
    Retorna (html, farol)."""
    diferenca = saipos_val - planilha_val
    farol, status = _classificar_farol(diferenca)
    cor_fundo, cor_texto = CORES_FAROL[farol]

    linhas_detalhe = ""
    for item in sorted(detalhe, key=lambda x: x["hora"]):
        linhas_detalhe += f"""
        <tr>
          <td style="padding:6px 10px;border:1px solid #eee;">{item['pedido']}</td>
          <td style="padding:6px 10px;border:1px solid #eee;">{item['hora']}</td>
          <td align="right" style="padding:6px 10px;border:1px solid #eee;">R$ {item['valor']:.2f}</td>
        </tr>
        """

    tabela_detalhe = ""
    if detalhe:
        tabela_detalhe = f"""
        <p style="font-size:12px;color:#888;margin:10px 0 4px;">
          Detalhe das vendas em {titulo.lower()} ({len(detalhe)} pedido(s)) — pedido e hora, para conferir na câmera se necessário:
        </p>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:12px;">
          <tr style="background:#F7F7F7;">
            <th align="left" style="padding:6px 10px;border:1px solid #eee;color:#777;">Pedido</th>
            <th align="left" style="padding:6px 10px;border:1px solid #eee;color:#777;">Hora</th>
            <th align="right" style="padding:6px 10px;border:1px solid #eee;color:#777;">Valor</th>
          </tr>
          {linhas_detalhe}
        </table>
        """

    bloco_deposito = ""
    if mostrar_deposito:
        bloco_deposito = f"""
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px;">
          <tr>
            <td style="background:#FFF4D6;border:1px solid #E8C547;border-radius:6px;padding:10px 14px;font-size:14px;color:#6B4E00;font-weight:700;">
              💰 Valor para depósito desse caixa: R$ {planilha_val:.2f}
            </td>
          </tr>
        </table>
        """

    html = f"""
    <h3 style="margin:18px 0 6px;color:#333;">{titulo}{" - Turno " + turno_label if turno_label else ""}</h3>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
      <tr style="background:#F0F0F0;">
        <th align="left" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Qtd. vendas</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Loja</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Saipos</th>
        <th align="center" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Status</th>
        <th align="right" style="padding:8px 10px;border:1px solid #ddd;color:#555;">Diferença</th>
      </tr>
      <tr>
        <td style="padding:8px 10px;border:1px solid #ddd;">{len(detalhe)}</td>
        <td align="right" style="padding:8px 10px;border:1px solid #ddd;">R$ {planilha_val:.2f}</td>
        <td align="right" style="padding:8px 10px;border:1px solid #ddd;">R$ {saipos_val:.2f}</td>
        <td align="center" style="padding:8px 10px;border:1px solid #ddd;background:{cor_fundo};color:{cor_texto};font-weight:700;">{status}</td>
        <td align="right" style="padding:8px 10px;border:1px solid #ddd;">R$ {diferenca:.2f}</td>
      </tr>
    </table>
    {bloco_deposito}
    {tabela_detalhe}
    """
    return html, farol


def montar_secao_dinheiro(turno_label, planilha_dinheiro, saipos_dinheiro_total, detalhe):
    return _montar_secao_com_detalhe("Dinheiro", turno_label, planilha_dinheiro, saipos_dinheiro_total, detalhe, mostrar_deposito=True)


def montar_secao_cortesia(turno_label, planilha_cortesia, saipos_cortesia_total, detalhe):
    return _montar_secao_com_detalhe("Cortesia", turno_label, planilha_cortesia, saipos_cortesia_total, detalhe, mostrar_deposito=False)


FAROL_EMOJI = {"verde": "🟢", "amarelo": "🟡", "vermelho": "🔴"}
FAROL_TEXTO = {"verde": "TUDO OK", "amarelo": "VERIFICAR", "vermelho": "URGENTE"}


def calcular_farol_geral(farois):
    if "vermelho" in farois:
        return "vermelho"
    if "amarelo" in farois:
        return "amarelo"
    if "verde" in farois:
        return "verde"
    return "cinza"


def enviar_email_conciliacao(token, data_str, secoes_html, farois_gerais, excecao_pix=None, modo_teste=False):
    ano_f, mes_f, dia_f = data_str.split("-")
    data_formatada = f"{dia_f}/{mes_f}/{ano_f}"

    farol_geral = calcular_farol_geral(farois_gerais)
    emoji_geral = FAROL_EMOJI.get(farol_geral, "⚪")
    texto_geral = FAROL_TEXTO.get(farol_geral, "SEM DADOS")

    faixa_teste = ""
    if modo_teste:
        faixa_teste = """
        <div style="background:#333;color:#fff;text-align:center;padding:8px;font-size:12px;font-weight:700;letter-spacing:1px;">
          ⚠️ E-MAIL DE TESTE — NÃO É O FECHAMENTO OFICIAL DO DIA
        </div>
        """

    secao_excecao = ""
    if excecao_pix:
        total_inter, total_pagseguro, diferenca = excecao_pix
        secao_excecao = f"""
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #F5D9A8;background:#FFF9E6;border-radius:6px;margin-top:16px;">
          <tr><td style="padding:12px 14px;font-size:12px;color:#666;">
            <b style="color:#8a6d1e;">Possível PIX direto no CNPJ:</b> o Banco Inter recebeu
            R$ {diferenca:.2f} a mais em PIX do que o registrado no PagSeguro
            (Inter R$ {total_inter:.2f} x PagSeguro R$ {total_pagseguro:.2f}).
            Confira e some manualmente ao fechamento, se confirmado.
          </td></tr>
        </table>
        """

    corpo_html = f"""
    <html><body style="margin:0;padding:0;background:#f2f2f2;font-family:'Segoe UI',Arial,sans-serif;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2f2;padding:24px 0;">
    <tr><td align="center">
    <table role="presentation" width="680" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,.08);">
      {faixa_teste}
      <tr>
        <td style="background:#6B0A0A;padding:16px 24px;">
          <table role="presentation" width="100%"><tr>
            <td style="color:#E8C547;font-size:18px;font-weight:800;">Fechamento de Caixa</td>
            <td align="right">
              <span style="background:#fff;color:#333;font-size:12px;font-weight:700;padding:5px 12px;border-radius:20px;">
                {emoji_geral} {texto_geral}
              </span>
            </td>
          </tr></table>
        </td>
      </tr>
      <tr>
        <td style="padding:16px 24px 4px;font-size:13px;">
          <b>Loja:</b> Japa Nobre — {CIDADE} &nbsp;|&nbsp; <b>Data:</b> {data_formatada}
        </td>
      </tr>
      <tr>
        <td style="padding:0 24px 20px;">
          {"".join(secoes_html)}
          {secao_excecao}
          <div style="border-top:1px solid #eee;margin-top:20px;padding-top:12px;font-size:11px;color:#aaa;text-align:center;">
            Gerado automaticamente pelo Portal Financeiro Japa Nobre
          </div>
        </td>
      </tr>
    </table>
    </td></tr>
    </table>
    </body></html>
    """

    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_REMETENTE_MAILBOX}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "message": {
            "subject": f"{'[TESTE] ' if modo_teste else ''}{emoji_geral} Fechamento Caixa {CIDADE} - {data_formatada}",
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": EMAIL_DESTINATARIO}}],
        }
    }
    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()


# ─────────────────────────────────────────────
# EXECUCAO PRINCIPAL
# ─────────────────────────────────────────────
def main():
    data_override = os.environ.get("DATA_ALVO", "").strip()
    if data_override:
        ontem = datetime.strptime(data_override, "%Y-%m-%d").date()
        print(f"Usando data manual (DATA_ALVO): {ontem}")
    else:
        ontem = date.today() - timedelta(days=1)
    data_str = ontem.strftime("%Y-%m-%d")
    print(f"Conciliando o dia: {data_str}")

    token = obter_token_graph()
    site_id = obter_site_id(token)
    item_id = localizar_arquivo_dashboard(token, site_id, ontem.month, ontem.year)
    linhas = buscar_linhas_do_dia(token, site_id, item_id, ontem)
    print(f"Turnos encontrados na planilha: {len(linhas)}")

    if not linhas:
        print("Nenhuma linha encontrada para esse dia na planilha. Encerrando.")
        return

    transacoes_hoje, validado_hoje = buscar_transacoes_pagseguro(data_str)
    print(f"Transacoes PagSeguro em {data_str}: {len(transacoes_hoje)} - validado={validado_hoje}")

    try:
        saipos_por_turno = montar_dados_saipos_por_turno(data_str)
        print(f"Turnos encontrados no Saipos: {list(saipos_por_turno.keys())}")
    except Exception as e:
        print(f"Falha ao buscar dados do Saipos: {e}")
        saipos_por_turno = {}

    secoes_html = []
    farois_gerais = []
    total_pix_dia_inteiro = 0.0

    def _vazio_turno():
        return {"categorias": {}, "dinheiro_detalhe": [], "cortesia_detalhe": []}

    if len(linhas) == 1:
        info = linhas[0]
        total_debito = somar_janela(transacoes_hoje, 0, 24 * 60, filtro=eh_debito)
        total_credito = somar_janela(transacoes_hoje, 0, 24 * 60, filtro=eh_credito)
        total_pix = somar_janela(transacoes_hoje, 0, 24 * 60, filtro=eh_pix)
        total_voucher_pg = somar_janela(transacoes_hoje, 0, 24 * 60, filtro=eh_voucher)
        total_pix_dia_inteiro = total_pix

        saipos_geral = _vazio_turno()
        for turno_dados in saipos_por_turno.values():
            for cat, val in turno_dados["categorias"].items():
                saipos_geral["categorias"][cat] = saipos_geral["categorias"].get(cat, 0.0) + val
            saipos_geral["dinheiro_detalhe"].extend(turno_dados["dinheiro_detalhe"])
            saipos_geral["cortesia_detalhe"].extend(turno_dados["cortesia_detalhe"])

        pagseguro_por_categoria = {"debito": total_debito, "credito": total_credito, "pix": total_pix, "voucher": total_voucher_pg}

        html, farois = montar_secao_turno(None, info["planilha"], saipos_geral["categorias"], pagseguro_por_categoria)
        secoes_html.append(html)
        farois_gerais.extend(farois)

        try:
            escrever_valor_online_planilha(token, site_id, item_id, info["linha"], saipos_geral["categorias"].get("online_parceiro", 0.0))
            print(f"Valor Online escrito na planilha (linha {info['linha']}): R$ {saipos_geral['categorias'].get('online_parceiro', 0.0):.2f}")
        except Exception as e:
            print(f"Falha ao escrever valor Online na planilha: {e}")

        html, farol = montar_secao_dinheiro(None, info["planilha"].get("dinheiro", 0.0), saipos_geral["categorias"].get("dinheiro", 0.0), saipos_geral["dinheiro_detalhe"])
        secoes_html.append(html)
        farois_gerais.append(farol)

        html, farol = montar_secao_cortesia(None, info["planilha"].get("cortesia", 0.0), saipos_geral["categorias"].get("cortesia", 0.0), saipos_geral["cortesia_detalhe"])
        secoes_html.append(html)
        farois_gerais.append(farol)

    else:
        dia_info, noite_info = linhas[0], linhas[1]
        hora_dia = dia_info["hora_fechamento"]
        hora_noite = noite_info["hora_fechamento"]

        total_debito_dia = total_credito_dia = total_pix_dia = total_voucher_dia = None
        total_debito_noite = total_credito_noite = total_pix_noite = total_voucher_noite = None

        if hora_dia:
            m_dia = parse_minutos(hora_dia)
            total_debito_dia = somar_janela(transacoes_hoje, 0, m_dia, filtro=eh_debito)
            total_credito_dia = somar_janela(transacoes_hoje, 0, m_dia, filtro=eh_credito)
            total_pix_dia = somar_janela(transacoes_hoje, 0, m_dia, filtro=eh_pix)
            total_voucher_dia = somar_janela(transacoes_hoje, 0, m_dia, filtro=eh_voucher)

        if hora_dia and hora_noite:
            m_dia = parse_minutos(hora_dia)
            m_noite = parse_minutos(hora_noite)
            cruzou_meia_noite = m_noite <= m_dia

            total_debito_noite = somar_janela(transacoes_hoje, m_dia, 24 * 60, filtro=eh_debito)
            total_credito_noite = somar_janela(transacoes_hoje, m_dia, 24 * 60, filtro=eh_credito)
            total_pix_noite = somar_janela(transacoes_hoje, m_dia, 24 * 60, filtro=eh_pix)
            total_voucher_noite = somar_janela(transacoes_hoje, m_dia, 24 * 60, filtro=eh_voucher)

            if cruzou_meia_noite:
                amanha_str = (ontem + timedelta(days=1)).strftime("%Y-%m-%d")
                print(f"Turno NOITE cruza a meia-noite - buscando tambem {amanha_str}")
                transacoes_amanha, _ = buscar_transacoes_pagseguro(amanha_str)
                total_debito_noite += somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_debito)
                total_credito_noite += somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_credito)
                total_pix_noite += somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_pix)
                total_voucher_noite += somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_voucher)

        total_pix_dia_inteiro = (total_pix_dia or 0) + (total_pix_noite or 0)

        saipos_dia = saipos_por_turno.get("Dia", _vazio_turno())
        saipos_noite = saipos_por_turno.get("Noite", _vazio_turno())

        html, farois = montar_secao_turno("DIA", dia_info["planilha"], saipos_dia["categorias"], {"debito": total_debito_dia, "credito": total_credito_dia, "pix": total_pix_dia, "voucher": total_voucher_dia})
        secoes_html.append(html)
        farois_gerais.extend(farois)

        try:
            escrever_valor_online_planilha(token, site_id, item_id, dia_info["linha"], saipos_dia["categorias"].get("online_parceiro", 0.0))
            print(f"Valor Online DIA escrito na planilha (linha {dia_info['linha']}): R$ {saipos_dia['categorias'].get('online_parceiro', 0.0):.2f}")
        except Exception as e:
            print(f"Falha ao escrever valor Online DIA na planilha: {e}")

        html, farol = montar_secao_dinheiro("DIA", dia_info["planilha"].get("dinheiro", 0.0), saipos_dia["categorias"].get("dinheiro", 0.0), saipos_dia["dinheiro_detalhe"])
        secoes_html.append(html)
        farois_gerais.append(farol)

        html, farol = montar_secao_cortesia("DIA", dia_info["planilha"].get("cortesia", 0.0), saipos_dia["categorias"].get("cortesia", 0.0), saipos_dia["cortesia_detalhe"])
        secoes_html.append(html)
        farois_gerais.append(farol)

        html, farois = montar_secao_turno("NOITE", noite_info["planilha"], saipos_noite["categorias"], {"debito": total_debito_noite, "credito": total_credito_noite, "pix": total_pix_noite, "voucher": total_voucher_noite})
        secoes_html.append(html)
        farois_gerais.extend(farois)

        try:
            escrever_valor_online_planilha(token, site_id, item_id, noite_info["linha"], saipos_noite["categorias"].get("online_parceiro", 0.0))
            print(f"Valor Online NOITE escrito na planilha (linha {noite_info['linha']}): R$ {saipos_noite['categorias'].get('online_parceiro', 0.0):.2f}")
        except Exception as e:
            print(f"Falha ao escrever valor Online NOITE na planilha: {e}")

        html, farol = montar_secao_dinheiro("NOITE", noite_info["planilha"].get("dinheiro", 0.0), saipos_noite["categorias"].get("dinheiro", 0.0), saipos_noite["dinheiro_detalhe"])
        secoes_html.append(html)
        farois_gerais.append(farol)

        html, farol = montar_secao_cortesia("NOITE", noite_info["planilha"].get("cortesia", 0.0), saipos_noite["categorias"].get("cortesia", 0.0), saipos_noite["cortesia_detalhe"])
        secoes_html.append(html)
        farois_gerais.append(farol)

    excecao = checar_excecao_pix_inter(data_str, total_pix_dia_inteiro)
    enviar_email_conciliacao(token, data_str, secoes_html, farois_gerais, excecao, modo_teste=bool(data_override))
    print("E-mail unico enviado com sucesso.")


if __name__ == "__main__":
    main()
