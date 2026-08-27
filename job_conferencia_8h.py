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
    """Busca as transacoes do PagSeguro EDI, com ate 5 tentativas e espera
    progressiva em caso de falha (mesmo padrao ja usado com o Saipos)."""
    url = f"https://edi.api.pagbank.com.br/movement/v3.00/transactional/{data_str}"
    params = {"pageNumber": 1, "pageSize": 1000}

    ultimo_erro = None
    for tentativa in range(5):
        try:
            resp = requests.get(url, params=params, auth=HTTPBasicAuth(PAGSEGURO_USER, PAGSEGURO_TOKEN), timeout=60)
            if resp.status_code == 200:
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
            ultimo_erro = f"Status {resp.status_code}: {resp.text[:200]}"
        except Exception as e:
            ultimo_erro = str(e)

        espera = 20 * (tentativa + 1)
        print(f"PagSeguro falhou (tentativa {tentativa + 1}/5): {ultimo_erro} — aguardando {espera}s")
        time.sleep(espera)

    raise RuntimeError(f"PagSeguro nao respondeu apos 5 tentativas. Ultimo erro: {ultimo_erro}")


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
    """Se transacoes for None (PagSeguro indisponivel apos as tentativas),
    retorna None - para o e-mail mostrar 'N/D', nunca um R$0,00 enganoso."""
    if transacoes is None:
        return None
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
    """Busca o PIX recebido no Banco Inter, com ate 5 tentativas e espera
    progressiva em caso de falha (mesmo padrao do PagSeguro e Saipos)."""
    ultimo_erro = None
    for tentativa in range(5):
        try:
            token = obter_token_inter()
            url = "https://cdpj.partners.bancointer.com.br/banking/v2/extrato"
            resp = requests.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                params={"dataInicio": data_str, "dataFim": data_str},
                cert=(INTER_CERT_PATH, INTER_KEY_PATH),
                timeout=60,
            )
            if resp.status_code == 200:
                transacoes = resp.json().get("transacoes", [])
                total_pix = 0.0
                for t in transacoes:
                    if t.get("tipoTransacao", "").upper() == "PIX" and t.get("tipoOperacao", "").upper() == "C":
                        total_pix += float(t.get("valor", 0))
                return total_pix
            ultimo_erro = f"Status {resp.status_code}: {resp.text[:200]}"
        except Exception as e:
            ultimo_erro = str(e)

        espera = 20 * (tentativa + 1)
        print(f"Banco Inter falhou (tentativa {tentativa + 1}/5): {ultimo_erro} — aguardando {espera}s")
        time.sleep(espera)

    raise RuntimeError(f"Banco Inter nao respondeu apos 5 tentativas. Ultimo erro: {ultimo_erro}")


def _extrair_pix_direto_declarado(obs_texto):
    """Extrai a declaração de PIX direto que o gerente registrou no
    formulário (formato 'PIXDIRETO:valor:motivo' dentro da Observação).
    Retorna (valor, motivo) ou (0.0, None) se não houver declaração."""
    if not obs_texto:
        return 0.0, None
    m = re.search(r"PIXDIRETO:([\d.]+):([^—]*)", obs_texto)
    if not m:
        return 0.0, None
    try:
        valor = float(m.group(1))
    except ValueError:
        valor = 0.0
    motivo = m.group(2).strip()
    return valor, motivo


def buscar_fotos_do_dia(token, site_id, data_str):
    """Busca os JSONs brutos salvos no Teams (pasta Recebimentos) para a
    data informada, e extrai as fotos (papel do Saipos + maquinas) de
    cada fechamento enviado naquele dia, prontas para anexar no e-mail."""
    ano, mes, _ = data_str.split("-")
    caminho = f"{ano}/{mes}/Recebimentos"
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{caminho}:/children"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        print(f"Nao encontrei a pasta de Recebimentos para {data_str}: {resp.status_code}")
        return []
    arquivos = resp.json().get("value", [])

    anexos = []
    for item in arquivos:
        nome = item.get("name", "")
        if not (nome.startswith(f"fechamento-{data_str}") and nome.endswith(".json")):
            continue

        download_url = item.get("@microsoft.graph.downloadUrl")
        if not download_url:
            continue
        r2 = requests.get(download_url)
        if r2.status_code != 200:
            continue
        payload = r2.json()

        protocolo = payload.get("protocolo", nome.replace(".json", ""))
        hora = payload.get("hora_fechamento", "")
        sufixo = f"{protocolo}" + (f"-{hora.replace(':', 'h')}" if hora else "")

        anexo_saipos = _montar_anexo_foto(f"saipos-{sufixo}.jpg", payload.get("foto_saipos_base64"))
        anexo_maquina = _montar_anexo_foto(f"maquinas-{sufixo}.jpg", payload.get("foto_maquina_base64"))
        if anexo_saipos:
            anexos.append(anexo_saipos)
        if anexo_maquina:
            anexos.append(anexo_maquina)

    return anexos


def _montar_anexo_foto(nome_arquivo, data_uri):
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


def checar_excecao_pix_inter(data_str, total_pix_pagseguro, observacoes_turnos=None):
    """Verifica se o Inter recebeu mais PIX do que o PagSeguro processou.
    Se o gerente ja declarou esse PIX direto no formulario (com valor
    batendo), trata como AVISO informativo. Se nao declarou (ou o valor
    nao bate), trata como SUSPEITA DE FRAUDE - alerta mais forte."""
    try:
        total_pix_inter = buscar_pix_recebido_inter(data_str)
    except Exception as e:
        print(f"Nao foi possivel checar excecoes no Banco Inter: {e}")
        return None

    diferenca = total_pix_inter - total_pix_pagseguro
    print(f"Checagem de excecao - Inter: R$ {total_pix_inter:.2f} | PagSeguro: R$ {total_pix_pagseguro:.2f}")

    if diferenca <= 5:
        return None

    # Soma tudo que o(s) gerente(s) declararam nos turnos do dia
    valor_declarado_total = 0.0
    motivos_declarados = []
    for obs in (observacoes_turnos or []):
        valor, motivo = _extrair_pix_direto_declarado(obs)
        if valor > 0:
            valor_declarado_total += valor
            if motivo:
                motivos_declarados.append(motivo)

    gerente_avisou_e_bate = abs(valor_declarado_total - diferenca) <= 5

    print(f"Possivel PIX direto no CNPJ detectado: R$ {diferenca:.2f} | Declarado pelo gerente: R$ {valor_declarado_total:.2f}")
    return (total_pix_inter, total_pix_pagseguro, diferenca, gerente_avisou_e_bate, valor_declarado_total, motivos_declarados)


# ─────────────────────────────────────────────
# SAIPOS
# ─────────────────────────────────────────────
def categorizar_pagamento_saipos(desc):
    """Classifica o texto livre do Saipos (desc_store_payment_type) em
    uma das categorias que usamos na conciliacao.

    IMPORTANTE: checa "ifood"/"99food" ANTES de checar "pix", porque o
    Saipos registra pagamentos feitos via PIX DENTRO do app do iFood
    como "Pix - iFood Pago" — esse dinheiro nao cai direto na nossa
    conta/maquininha, fica retido com o parceiro e eh repassado junto
    com o resto do "Online/Parceiros". So o PIX pago diretamente pelo
    cliente (presencial/balcao) deve contar como "pix" de verdade."""
    d = (desc or "").lower()
    if "ifood" in d or "99food" in d or "99 food" in d:
        return "online_parceiro"
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
    - pix_ifood_total: quanto do total "online_parceiro" veio de PIX pago
      dentro do app do iFood (so para exibir uma nota explicativa, ja que
      esse valor conta como "online_parceiro", nao como "pix")
    """
    vendas = buscar_vendas_saipos(data_str)

    dados = {}
    for v in vendas:
        turno = (v.get("store_shift") or {}).get("desc_store_shift", "Desconhecido")
        if turno not in dados:
            dados[turno] = {"categorias": {}, "dinheiro_detalhe": [], "cortesia_detalhe": [], "pix_ifood_total": 0.0}

        for p in (v.get("payments") or []):
            desc = p.get("desc_store_payment_type", "")
            valor = float(p.get("payment_amount", 0))
            categoria = categorizar_pagamento_saipos(desc)

            dados[turno]["categorias"][categoria] = dados[turno]["categorias"].get(categoria, 0.0) + valor

            desc_lower = desc.lower()
            if "pix" in desc_lower and ("ifood" in desc_lower or "99food" in desc_lower or "99 food" in desc_lower):
                dados[turno]["pix_ifood_total"] += valor

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


def _linha_tabela(fonte, saipos_val, planilha_val, pagseguro_val=None, informativo=False, saipos_val_consolidado=None):
    """Monta uma linha <tr> da tabela principal. pagseguro_val=None significa
    'nao aplicavel' (mostra traco). Se informativo=True, a linha nao entra no
    calculo do farol (ex: Online/Parceiros, que a loja sempre lanca como 0
    porque nao tem acesso a esse valor - quem preenche depois e o escritorio).

    saipos_val_consolidado: se informado, usa ESSE valor (nao o saipos_val
    exibido na coluna) para calcular status/diferenca. Serve para casos
    como PIX, onde parte do valor esta "escondida" em outra categoria
    (ex: PIX pago via iFood conta como Online) - o que importa para o
    farol e' se o TOTAL consolidado bate, mesmo que a divisao por
    categoria não bata exatamente.

    Retorna (html, farol_ou_none)."""
    valor_para_calculo = saipos_val_consolidado if saipos_val_consolidado is not None else saipos_val

    if informativo:
        cor_fundo, cor_texto = CORES_FAROL["cinza"]
        status = "Preenchido auto."
        diferenca_txt = "—"
        farol = None
    elif valor_para_calculo is None or planilha_val is None:
        cor_fundo, cor_texto = CORES_FAROL["cinza"]
        status = "N/D"
        diferenca_txt = "—"
        farol = None
    else:
        diferenca = valor_para_calculo - planilha_val
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


def _detectar_trocas_categoria(deltas):
    """Procura pares de categorias cujas diferencas se cancelam (uma sobrou,
    outra faltou, valores parecidos) - sinal de que o gerente lancou o
    valor na categoria errada. Retorna (lista_de_notas, nomes_ja_explicados)."""
    notas = []
    explicados = set()
    nomes = list(deltas.keys())

    for i in range(len(nomes)):
        for j in range(i + 1, len(nomes)):
            nome_a, nome_b = nomes[i], nomes[j]
            if nome_a in explicados or nome_b in explicados:
                continue
            delta_a = deltas[nome_a]["delta"]
            delta_b = deltas[nome_b]["delta"]
            # Uma sobrou (positivo) e a outra faltou (negativo), em valores parecidos
            if delta_a * delta_b < 0 and abs(delta_a + delta_b) <= 5:
                info_a = deltas[nome_a]
                info_b = deltas[nome_b]
                quem_faltou = nome_a if delta_a < 0 else nome_b
                quem_sobrou = nome_b if delta_a < 0 else nome_a
                valor_movido = abs(delta_a)
                info_falta = deltas[quem_faltou]
                info_sobra = deltas[quem_sobrou]
                notas.append(
                    f"<li>O gerente bateu R$ {info_falta['loja']:.2f} no <b>{quem_faltou}</b>, mas o Saipos indica "
                    f"R$ {info_falta['saipos']:.2f} vendido (faltou R$ {valor_movido:.2f}). Ao mesmo tempo, o "
                    f"<b>{quem_sobrou}</b> teve R$ {valor_movido:.2f} a mais do que o Saipos indicava "
                    f"(Loja R$ {info_sobra['loja']:.2f} vs Saipos R$ {info_sobra['saipos']:.2f}). "
                    f"Isso sugere que esse valor foi lançado na categoria errada — o consolidado do dia bateu, "
                    f"mas vale alinhar com o gerente para não repetir.</li>"
                )
                explicados.add(nome_a)
                explicados.add(nome_b)

    return notas, explicados


def montar_secao_turno(turno_label, planilha, saipos_categorias, pagseguro_por_categoria, pix_ifood_total=0.0):
    """Monta as linhas da tabela principal para um turno (exclui Dinheiro e
    Cortesia, que tem secao propria com relatorio detalhado).

    O farol GERAL desta secao (usado no farol do e-mail) e' baseado no
    CONSOLIDADO (soma de todas as categorias) - se o total do turno bate,
    o farol fica verde mesmo que categorias individuais estejam diferentes
    entre si (ex: sobrou no Debito, faltou no Credito). Cada linha da
    tabela continua mostrando a comparacao real por categoria, para
    transparencia, mas isso e' informativo - o que decide o farol e' o
    total consolidado.

    Retorna (html, lista_de_farois) - normalmente 1 farol (o do total
    consolidado do turno)."""
    linhas_html = []

    mapeamentos = [
        ("Débito", "debito", "debito", "debito", False),
        ("Crédito", "credito", "credito", "credito", False),
        ("PIX", "pix", "pix", "pix", False),
        ("Voucher/Vale Refeição", "voucher_vale", "vale_voucher_soma", "voucher", False),
        ("Online/Parceiros (iFood, 99Food etc.)", "online_parceiro", "online", None, True),
    ]

    deltas = {}  # nome -> {"loja": x, "saipos": y, "delta": saipos-loja} - so categorias que contam no consolidado

    for fonte_label, cat_saipos, cat_planilha, cat_pagseguro, informativo in mapeamentos:
        saipos_val = saipos_categorias.get(cat_saipos, 0.0)

        if cat_planilha == "vale_voucher_soma":
            planilha_val = planilha.get("vale", 0.0) + planilha.get("voucher", 0.0)
        else:
            planilha_val = planilha.get(cat_planilha, 0.0)

        pagseguro_val = pagseguro_por_categoria.get(cat_pagseguro) if cat_pagseguro else None

        # PIX usa o valor consolidado (direto + iFood) so para exibicao do
        # status individual da linha - mas o que decide o farol GERAL e'
        # sempre o total do turno (calculado abaixo), entao aqui only
        # mostra a linha corretamente.
        saipos_val_para_linha = saipos_val
        if fonte_label == "PIX" and pix_ifood_total > 0:
            saipos_val_para_linha = saipos_val + pix_ifood_total

        html, _ = _linha_tabela(fonte_label, saipos_val, planilha_val, pagseguro_val, informativo=informativo, saipos_val_consolidado=(saipos_val_para_linha if fonte_label == "PIX" else None))
        linhas_html.append(html)

        if not informativo:
            deltas[fonte_label] = {
                "loja": planilha_val,
                "saipos": saipos_val_para_linha,
                "delta": saipos_val_para_linha - planilha_val,
            }

        # Nota explicativa logo abaixo da linha de PIX, se houver PIX pago
        # dentro do app do iFood (esse valor conta em "Online", nao aqui).
        if fonte_label == "PIX" and pix_ifood_total > 0:
            pix_direto = saipos_val
            soma_total = pix_direto + pix_ifood_total
            linhas_html.append(f"""
            <tr>
              <td colspan="6" style="padding:6px 10px;border:1px solid #ddd;background:#FAFAFA;font-size:11px;color:#777;">
                📌 Nota: R$ {pix_ifood_total:.2f} desse PIX foi pago <b>dentro do app do iFood</b>
                (já contabilizado na linha "Online/Parceiros" abaixo, não cai direto na conta).
                O status acima já considera o total consolidado
                (PIX direto R$ {pix_direto:.2f} + PIX via iFood R$ {pix_ifood_total:.2f} = R$ {soma_total:.2f}).
              </td>
            </tr>
            """)

    # ── Calcula o CONSOLIDADO do turno (soma de todas as categorias que contam) ──
    total_loja = sum(d["loja"] for d in deltas.values())
    total_saipos = sum(d["saipos"] for d in deltas.values())
    diferenca_total = total_saipos - total_loja

    # ── Loja parece nao ter enviado o fechamento ainda (tudo zerado, mas
    # o Saipos mostra movimento real) - isso NAO e' um erro de conferencia,
    # e' so falta de dado. Tratar diferente de uma divergencia de verdade.
    fechamento_nao_enviado = total_loja == 0 and total_saipos > 5

    if fechamento_nao_enviado:
        farol_geral, status_geral = "cinza", "Dados não preenchidos"
        aviso_nao_enviado = f"""
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:2px solid #999;background:#F0F0F0;border-radius:6px;margin-bottom:10px;">
          <tr><td style="padding:12px 14px;font-size:13px;color:#444;">
            <b>⚪ A LOJA NÃO PREENCHEU OS DADOS deste turno.</b><br>
            O Saipos já mostra R$ {total_saipos:.2f} em vendas, mas o formulário de fechamento
            não foi preenchido — por isso os valores abaixo NÃO representam uma comparação real,
            e as diferenças por categoria não são erros de caixa. Assim que a loja preencher o
            fechamento, rode a conciliação de novo para ver os valores reais comparados.
          </td></tr>
        </table>
        """
    else:
        farol_geral, status_geral = _classificar_farol(diferenca_total)
        aviso_nao_enviado = ""

    # ── Detecta trocas de categoria (sobrou em uma, faltou em outra) ──
    notas_trocas, categorias_explicadas = _detectar_trocas_categoria(deltas)

    # ── Notas para diferencas que sobraram sem explicacao automatica ──
    # (nao gera essas notas se o motivo ja e' "fechamento nao enviado")
    notas_pendentes = []
    if not fechamento_nao_enviado:
        for nome, info in deltas.items():
            if nome in categorias_explicadas:
                continue
            if nome == "PIX" and pix_ifood_total > 0:
                continue  # ja explicado na nota do PIX-iFood acima
            if abs(info["delta"]) > 5:
                sinal = "a mais" if info["delta"] > 0 else "a menos"
                notas_pendentes.append(
                    f"<li><b>{nome}</b> teve uma diferença de R$ {abs(info['delta']):.2f} ({sinal}) que não "
                    f"foi automaticamente explicada — vale conferir com o gerente o motivo específico.</li>"
                )

    todas_notas = notas_trocas + notas_pendentes
    bloco_notas = ""
    if todas_notas:
        bloco_notas = f"""
        <div style="margin-top:8px;padding:10px 14px;background:#FFF9E6;border-left:3px solid #E8C547;border-radius:4px;font-size:12px;color:#555;">
          <b>🔍 Observações para o conferente:</b>
          <ul style="margin:6px 0 0;padding-left:18px;">
            {"".join(todas_notas)}
          </ul>
        </div>
        """

    nota_online = """
    <p style="font-size:11px;color:#999;margin:4px 0 0;">
      * Online/Parceiros: a loja não tem acesso a esse valor no fechamento (sempre lança 0).
      O valor mostrado na coluna Saipos acima já foi preenchido automaticamente na planilha
      Dashboard pelo sistema. Essa linha não entra no status/farol do fechamento.
    </p>
    """

    resumo_consolidado = f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px;">
      <tr>
        <td style="background:{CORES_FAROL[farol_geral][0]};border-radius:6px;padding:8px 12px;font-size:12px;color:{CORES_FAROL[farol_geral][1]};font-weight:700;">
          Consolidado do turno: Loja R$ {total_loja:.2f} | Saipos R$ {total_saipos:.2f} | Diferença R$ {diferenca_total:.2f} — {status_geral}
        </td>
      </tr>
    </table>
    """

    tabela = f"""
    <h3 style="margin:18px 0 6px;color:#333;">{"Turno " + turno_label if turno_label else "Fechamento do dia"}</h3>
    {aviso_nao_enviado}
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
    {resumo_consolidado}
    {nota_online}
    {bloco_notas}
    """
    return tabela, [farol_geral]


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


def enviar_email_conciliacao(token, data_str, secoes_html, farois_gerais, excecao_pix=None, anexos=None):
    ano_f, mes_f, dia_f = data_str.split("-")
    data_formatada = f"{dia_f}/{mes_f}/{ano_f}"

    farol_geral = calcular_farol_geral(farois_gerais)
    emoji_geral = FAROL_EMOJI.get(farol_geral, "⚪")
    texto_geral = FAROL_TEXTO.get(farol_geral, "SEM DADOS")

    secao_excecao = ""
    if excecao_pix:
        total_inter, total_pagseguro, diferenca, gerente_avisou_e_bate, valor_declarado, motivos = excecao_pix

        if gerente_avisou_e_bate:
            motivos_txt = "; ".join(motivos) if motivos else "não informado"
            secao_excecao = f"""
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #B8E0C8;background:#EAF7EF;border-radius:6px;margin-top:16px;">
              <tr><td style="padding:12px 14px;font-size:12px;color:#666;">
                <b style="color:#1e7a45;">✅ PIX direto no CNPJ — já informado pelo gerente:</b>
                R$ {diferenca:.2f} recebidos direto na conta, conforme declarado no fechamento
                (motivo: {motivos_txt}). Valor bate com o detectado no Banco Inter — sem necessidade
                de ação adicional.
              </td></tr>
            </table>
            """
        else:
            aviso_declarado = ""
            if valor_declarado > 0:
                aviso_declarado = f" (o gerente declarou R$ {valor_declarado:.2f}, mas o valor não bate com o detectado — confira)"
            secao_excecao = f"""
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:2px solid #E67E22;background:#FDF3E7;border-radius:6px;margin-top:16px;">
              <tr><td style="padding:14px 16px;font-size:13px;color:#7a4a12;">
                <b style="font-size:14px;">⚠️ PIX recebido no Banco Inter sem explicação no fechamento</b><br><br>
                O Banco Inter recebeu <b>R$ {diferenca:.2f}</b> a mais em PIX do que o registrado no
                PagSeguro (Inter R$ {total_inter:.2f} x PagSeguro R$ {total_pagseguro:.2f}), e o gerente
                <b>não declarou</b> esse recebimento no fechamento{aviso_declarado}.
                <br><br>
                <b>Causa mais provável:</b> uso da maquininha <b>InterPag</b> (que não é integrada ao sistema
                e deveria ser usada só em emergências). Também pode ser PIX direto de um cliente.
                <br><br>
                <b>Ação recomendada:</b> confirmar com o gerente da loja qual foi o motivo, e reforçar que
                o fechamento deve informar sempre que a InterPag ou PIX direto forem usados, no campo
                específico do formulário.
              </td></tr>
            </table>
            """

    corpo_html = f"""
    <html><body style="margin:0;padding:0;background:#f2f2f2;font-family:'Segoe UI',Arial,sans-serif;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2f2;padding:24px 0;">
    <tr><td align="center">
    <table role="presentation" width="680" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,.08);">
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
            "subject": f"{emoji_geral} Fechamento Caixa {CIDADE} - {data_formatada}",
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": EMAIL_DESTINATARIO}}],
            "attachments": anexos or [],
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
        print("Nenhuma linha encontrada para esse dia na planilha - conciliacao incompleta.")
        print("Mesmo assim, tentando enviar as fotos do dia (independente da conciliacao).")
        try:
            anexos_fotos = buscar_fotos_do_dia(token, site_id, data_str)
        except Exception as e:
            print(f"Falha ao buscar fotos do dia: {e}")
            anexos_fotos = []

        secoes_html = ["""
        <p style="color:#c00;font-size:13px;">
          ⚠️ A linha desse dia ainda não foi encontrada na planilha Dashboard —
          não foi possível fazer a conciliação completa (Débito/Crédito/PIX/Saipos).
          As fotos anexadas abaixo são do fechamento enviado pela loja, para conferência manual.
        </p>
        """]
        enviar_email_conciliacao(token, data_str, secoes_html, ["cinza"], None, anexos=anexos_fotos)
        return

    try:
        transacoes_hoje, validado_hoje = buscar_transacoes_pagseguro(data_str)
    except Exception as e:
        print(f"PagSeguro indisponivel apos todas as tentativas: {e}")
        transacoes_hoje, validado_hoje = None, False
    print(f"Transacoes PagSeguro em {data_str}: {len(transacoes_hoje) if transacoes_hoje is not None else 'INDISPONIVEL'} - validado={validado_hoje}")

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
        return {"categorias": {}, "dinheiro_detalhe": [], "cortesia_detalhe": [], "pix_ifood_total": 0.0}

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
            saipos_geral["pix_ifood_total"] += turno_dados.get("pix_ifood_total", 0.0)

        pagseguro_por_categoria = {"debito": total_debito, "credito": total_credito, "pix": total_pix, "voucher": total_voucher_pg}

        html, farois = montar_secao_turno(None, info["planilha"], saipos_geral["categorias"], pagseguro_por_categoria, saipos_geral["pix_ifood_total"])
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

        observacoes_do_dia = [info.get("observacao", "")]

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
                try:
                    transacoes_amanha, _ = buscar_transacoes_pagseguro(amanha_str)
                except Exception as e:
                    print(f"PagSeguro (dia seguinte) indisponivel apos todas as tentativas: {e}")
                    transacoes_amanha = None

                extra_debito = somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_debito)
                extra_credito = somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_credito)
                extra_pix = somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_pix)
                extra_voucher = somar_janela(transacoes_amanha, 0, m_noite, filtro=eh_voucher)

                # Se qualquer um dos dois lados for None (falha), o total vira None
                # (N/D) - nao da pra "somar zero escondido" com um valor real.
                total_debito_noite = None if (total_debito_noite is None or extra_debito is None) else total_debito_noite + extra_debito
                total_credito_noite = None if (total_credito_noite is None or extra_credito is None) else total_credito_noite + extra_credito
                total_pix_noite = None if (total_pix_noite is None or extra_pix is None) else total_pix_noite + extra_pix
                total_voucher_noite = None if (total_voucher_noite is None or extra_voucher is None) else total_voucher_noite + extra_voucher

        total_pix_dia_inteiro = (total_pix_dia or 0) + (total_pix_noite or 0)

        saipos_dia = saipos_por_turno.get("Dia", _vazio_turno())
        saipos_noite = saipos_por_turno.get("Noite", _vazio_turno())

        html, farois = montar_secao_turno("DIA", dia_info["planilha"], saipos_dia["categorias"], {"debito": total_debito_dia, "credito": total_credito_dia, "pix": total_pix_dia, "voucher": total_voucher_dia}, saipos_dia.get("pix_ifood_total", 0.0))
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

        html, farois = montar_secao_turno("NOITE", noite_info["planilha"], saipos_noite["categorias"], {"debito": total_debito_noite, "credito": total_credito_noite, "pix": total_pix_noite, "voucher": total_voucher_noite}, saipos_noite.get("pix_ifood_total", 0.0))
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

        observacoes_do_dia = [dia_info.get("observacao", ""), noite_info.get("observacao", "")]

    excecao = checar_excecao_pix_inter(data_str, total_pix_dia_inteiro, observacoes_do_dia)

    # O farol GERAL so pode fechar "tudo OK" se a excecao do Inter (quando
    # existe) ja estiver explicada. Se o Banco Inter recebeu PIX que nao
    # bate com o PagSeguro e o gerente NAO declarou o motivo, isso conta
    # como pendencia - o farol geral nao fecha verde ate isso ser resolvido.
    if excecao:
        _, _, _, gerente_avisou_e_bate, _, _ = excecao
        if not gerente_avisou_e_bate:
            farois_gerais.append("amarelo")

    try:
        anexos_fotos = buscar_fotos_do_dia(token, site_id, data_str)
        print(f"Fotos encontradas para anexar: {len(anexos_fotos)}")
    except Exception as e:
        print(f"Falha ao buscar fotos do dia: {e}")
        anexos_fotos = []

    enviar_email_conciliacao(token, data_str, secoes_html, farois_gerais, excecao, anexos=anexos_fotos)
    print("E-mail unico enviado com sucesso.")


if __name__ == "__main__":
    main()
