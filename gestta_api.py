# -*- coding: utf-8 -*-
"""
gestta_api.py
-------------
Camada de acesso a API do Gestta para o ONE_API.

Portado do projeto Libby (canella_libby), onde o padrao ja esta validado em
producao:
  - contatos/gestta_auth.py     -> resolucao de token
  - contatos/messenger_onvio.py -> listagem/indexacao de contatos
  - avisos_vencimento/avisos_vencimento_gestta.py -> leitura do token do Chrome

DOIS TOKENS, NAO UM
-------------------
Essa e a pegadinha que o Libby documentou e custa 401 se esquecida:
  * /core/*            -> token CORE      (localStorage 'ngStorage-jwt')
  * /messenger-admin/* -> token MESSENGER (localStorage 'user-jwt')
O ONE usa o Messenger, entao o token relevante aqui e o 'user-jwt'.

ESCOPO
------
Auth, contatos e ENVIO (texto, arquivo, grupo, desconsiderar). O contrato de
envio esta documentado em API_GESTTA_MESSENGER.md, descoberto por captura do
trafego real da tela do chat.
"""

import json
import os
import re
import time
import base64
import uuid as _uuid
import mimetypes
import unicodedata

try:
    import requests
except ImportError:
    requests = None

# ------------------------------------------------------------------ config ---

API_BASE = "https://api.gestta.com.br"

# Duas rotas de contato, com respostas DIFERENTES:
#  - /attendance-core/*  -> lista ARRAY, inclui grupos (is_group, group_id).
#    E a que a tela do chat usa; e a unica que traz o group_id necessario para
#    enviar em grupo. Preferida.
#  - /messenger-admin/*  -> paginada ({docs, hasNextPage}), so contatos, sem
#    group_id. E a que o Libby usa (messenger_onvio.py). Fallback.
CONTATOS_URL = API_BASE + "/attendance-core/company/contact"
CONTATOS_URL_ADMIN = API_BASE + "/messenger-admin/company/contact"
# O ONE nao abre o Chrome com --remote-debugging-port, entao nao da para anexar
# numa porta: lemos o token abrindo o Chrome no MESMO perfil (a sessao vem dele).
# CHROME_DEBUG so serve se houver um Chrome aberto manualmente com a flag.
CHROME_DEBUG = os.environ.get("GESTTA_CHROME_DEBUG", "")
PERFIL_DIR = r"C:\PerfisChrome\automacao_perfil%s"
PERFIL_PADRAO = os.environ.get("ONE_PERFIL_CHROME", "1")
URL_CHAT = "https://app.gestta.com.br/attendance/#/chat/contact-list"
PAGE_LIMIT = 200
MARGEM_SEG = 3600          # nao usa token com menos de 1h de validade

# Id da empresa (Canella & Santos) exigido no multipart do upload.
COMPANY_ID = os.environ.get("GESTTA_COMPANY_ID", "6418e03fe478c70006b05005")

TIMEOUT = 60

# Cache em memoria do token, para nao reabrir o Chrome a cada chamada.
_token_cache = {"messenger": None, "core": None}

# Chave do localStorage por tipo de token (ver docstring).
_LS_KEY = {"messenger": "user-jwt", "core": "ngStorage-jwt"}
_ENV_KEY = {"messenger": "MESSENGER_TOKEN", "core": "GESTTA_TOKEN"}


# ------------------------------------------------------------------ token ----

def _normalizar(tok):
    """Garante o prefixo 'JWT ' que a API espera no header authorization."""
    if not tok:
        return None
    tok = str(tok).strip().strip('"')
    return tok if tok.upper().startswith("JWT ") else "JWT " + tok


def _expiracao(tok):
    """Timestamp 'exp' do JWT (0 se nao decodificar)."""
    try:
        payload = tok.replace("JWT ", "").replace("jwt ", "").split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return int(json.loads(base64.urlsafe_b64decode(payload)).get("exp") or 0)
    except Exception:
        return 0


def token_valido(tok):
    """True se o token ainda tem folga suficiente de validade."""
    return bool(tok) and (_expiracao(tok) - time.time()) > MARGEM_SEG


def _ler_token_do_driver(drv, tipo):
    """Le a chave do localStorage numa aba do Gestta ja aberta no driver."""
    chave = _LS_KEY[tipo]
    script = ("try{return JSON.parse(localStorage.getItem('%s'));}"
              "catch(e){return localStorage.getItem('%s');}" % (chave, chave))
    for h in drv.window_handles:
        drv.switch_to.window(h)
        if "gestta.com.br" in (drv.current_url or ""):
            tok = drv.execute_script(script)
            if tok:
                return _normalizar(tok)
    return None


def ler_token_do_driver(drv, tipo="messenger", log=print):
    """Le o token de um driver que o CHAMADOR ja controla.

    Este e o caminho preferido dentro do ONE: o app ja tem o Chrome aberto e
    logado, entao aproveitamos a sessao sem abrir um segundo navegador.
    """
    try:
        tok = _ler_token_do_driver(drv, tipo)
    except Exception as e:
        log("Token %s: falha ao ler do driver -> %s" % (tipo, str(e)[:150]))
        return None
    if tok:
        _token_cache[tipo] = tok
    return tok


def _ler_do_chrome(tipo, perfil=None, log=print):
    """Le o JWT abrindo o Chrome no perfil do ONE (a sessao logada vem dele).

    Usado fora do app (scripts/diagnostico). Se GESTTA_CHROME_DEBUG estiver
    definido, anexa naquela porta em vez de abrir um Chrome novo.
    """
    try:
        from selenium import webdriver
    except ImportError:
        log("Selenium nao disponivel para ler o token.")
        return None

    opts = webdriver.ChromeOptions()
    proprio = not CHROME_DEBUG
    if CHROME_DEBUG:
        opts.add_experimental_option("debuggerAddress", CHROME_DEBUG)
    else:
        user_data_dir = PERFIL_DIR % (perfil or PERFIL_PADRAO)
        if not os.path.isdir(user_data_dir):
            log("Token %s: perfil nao encontrado (%s)." % (tipo, user_data_dir))
            return None
        opts.add_argument("--user-data-dir=%s" % user_data_dir)
        opts.add_argument("--headless=new")
        opts.add_argument("--lang=pt-BR")

    drv = None
    try:
        drv = webdriver.Chrome(options=opts)
        tok = _ler_token_do_driver(drv, tipo)
        if tok:
            return tok
        if proprio:
            # Perfil logado, mas sem aba do Gestta: abre para popular o storage.
            drv.get(URL_CHAT)
            for _ in range(20):
                time.sleep(1)
                tok = _ler_token_do_driver(drv, tipo)
                if tok:
                    return tok
        log("Token %s: nao encontrei o '%s' na sessao." % (tipo, _LS_KEY[tipo]))
        return None
    except Exception as e:
        msg = str(e)
        if "user data directory is already in use" in msg.lower():
            log("Token %s: o perfil esta em uso. Feche o Chrome do ONE, ou passe "
                "o token em %s." % (tipo, _ENV_KEY[tipo]))
        else:
            log("Token %s: falha ao ler do Chrome -> %s" % (tipo, msg[:150]))
        return None
    finally:
        if drv is not None:
            try:
                drv.quit() if proprio else drv.service.stop()
            except Exception:
                pass


def obter_token(tipo="messenger", perfil=None, log=print):
    """Resolve o token na ordem: cache -> env -> Chrome logado.

    `tipo`: 'messenger' (rotas de chat/envio) ou 'core' (/core/*).
    Dentro do ONE, prefira ler_token_do_driver(driver) — aproveita o Chrome
    que o app ja tem aberto, sem abrir um segundo.
    Levanta RuntimeError se nao conseguir.
    """
    if tipo not in _LS_KEY:
        raise ValueError("tipo deve ser 'messenger' ou 'core'")

    em_cache = _token_cache.get(tipo)
    if token_valido(em_cache):
        return em_cache

    do_env = _normalizar(os.environ.get(_ENV_KEY[tipo]))
    if token_valido(do_env):
        _token_cache[tipo] = do_env
        return do_env

    do_chrome = _ler_do_chrome(tipo, perfil=perfil, log=log)
    if do_chrome:
        _token_cache[tipo] = do_chrome
        if not token_valido(do_chrome):
            log("Aviso: token %s lido, mas perto de expirar." % tipo)
        return do_chrome

    raise RuntimeError(
        "Token %s indisponivel. Abra o Chrome de automacao e faca login no "
        "Gestta, ou defina a variavel de ambiente %s." % (tipo, _ENV_KEY[tipo])
    )


def limpar_cache():
    """Forca releitura do token na proxima chamada (util apos 401)."""
    _token_cache["messenger"] = None
    _token_cache["core"] = None


def _headers(tok):
    return {"authorization": tok,
            "Accept": "application/json",
            "Content-Type": "application/json"}


# --------------------------------------------------------------- contatos ----

def _norm_code(code):
    """Normaliza codigo para casamento: '09' -> '9', '55.0' -> '55'."""
    s = str(code or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = s.lstrip("0")
    return s if s else "0"


def extrair_codigos(name):
    """Codigos de empresa embutidos no inicio do nome do contato.

    Ex.: '09/55/884 - JOSIAS' -> ['9', '55', '884']
    """
    if not name:
        return []
    seg = re.split(r"\s*-\s*[A-Za-zÀ-ÿ]", name, maxsplit=1)[0]
    vistos, out = set(), []
    for g in re.findall(r"\d+", seg):
        c = _norm_code(g)
        if c not in vistos:
            vistos.add(c)
            out.append(c)
    return out


def _nome_pessoa(name):
    """Parte do nome apos os codigos (o nome da pessoa)."""
    if not name:
        return ""
    partes = re.split(r"\s*-\s*", name, maxsplit=1)
    return partes[1].strip() if len(partes) > 1 else name.strip()


def formatar_telefone(numero):
    """Formata como '55 24 98189-4647'. Devolve o original se nao reconhecer."""
    original = str(numero or "").strip()
    d = re.sub(r"\D", "", original)
    if not d:
        return ""
    if len(d) in (10, 11):          # sem codigo do pais
        d = "55" + d
    if d.startswith("55") and len(d) == 13:     # celular
        return "55 %s %s-%s" % (d[2:4], d[4:9], d[9:])
    if d.startswith("55") and len(d) == 12:     # fixo
        return "55 %s %s-%s" % (d[2:4], d[4:8], d[8:])
    return original


def listar_contatos(token=None, session=None, log=print):
    """Gera contatos E grupos do Messenger.

    Usa /attendance-core/company/contact (array unico, com is_group/group_id).
    Se falhar, cai no /messenger-admin/* paginado do Libby — que serve para
    contatos individuais, mas nao traz group_id.
    """
    if requests is None:
        raise RuntimeError("A biblioteca 'requests' e necessaria (pip install requests).")
    token = token or obter_token("messenger", log=log)
    s = session or requests.Session()

    try:
        r = s.get(CONTATOS_URL, headers=_headers(token), timeout=TIMEOUT)
        if r.status_code == 401:
            limpar_cache()
            raise RuntimeError("401 no Messenger: token invalido/expirado. "
                               "Confirme que e o 'user-jwt', nao o token core.")
        r.raise_for_status()
        dados = r.json()
        if isinstance(dados, list):
            for d in dados:
                yield d
            return
        # Resposta inesperada: tenta o formato paginado abaixo.
        log("Contatos: resposta nao-lista em attendance-core; tentando messenger-admin.")
    except RuntimeError:
        raise
    except Exception as e:
        log("Contatos: attendance-core falhou (%s); tentando messenger-admin."
            % str(e)[:120])

    page = 1
    while True:
        r = s.get(CONTATOS_URL_ADMIN, headers=_headers(token),
                  params={"page": page, "limit": PAGE_LIMIT, "sort": "name"},
                  timeout=TIMEOUT)
        if r.status_code == 401:
            limpar_cache()
            raise RuntimeError("401 no Messenger: token invalido/expirado.")
        r.raise_for_status()
        j = r.json()
        for d in j.get("docs", []):
            yield d
        if not j.get("hasNextPage"):
            break
        page += 1
        if page > 100:      # trava de seguranca
            break


def _info_contato(c):
    """Normaliza um registro da API para o formato usado no envio."""
    name = c.get("name", "") or ""
    return {
        "nome": _nome_pessoa(name),
        "telefone": formatar_telefone(c.get("phone_number", "")),
        "raw_name": name,
        "id": c.get("_id"),                       # company_contact das rotas de envio
        "is_group": bool(c.get("is_group")),
        "group_id": c.get("group_id") or None,    # obrigatorio para enviar em grupo
    }


def indexar_por_codigo(token=None, session=None, log=print, preferir_grupo=False):
    """Indice {codigo_empresa: info} a partir de UMA varredura da API.

    `info` traz id (company_contact), telefone, is_group e group_id — tudo que
    enviar() precisa. Quando o mesmo codigo aparece em contato e grupo,
    `preferir_grupo` decide qual vence; por padrao vence o contato individual,
    espelhando a ordem do ONE (usa grupo so quando o contato e "NONE").
    """
    indice = {}
    total = grupos = 0
    for c in listar_contatos(token=token, session=session, log=log):
        total += 1
        info = _info_contato(c)
        if info["is_group"]:
            grupos += 1
        for cod in extrair_codigos(info["raw_name"]):
            atual = indice.get(cod)
            if atual is None:
                indice[cod] = info
            elif preferir_grupo and info["is_group"] and not atual["is_group"]:
                indice[cod] = info
    log("Messenger: %d registros lidos (%d grupos), %d codigos indexados."
        % (total, grupos, len(indice)))
    return indice


def _chave_nome(nome):
    """Normaliza um nome para casamento tolerante.

    O Gestta guarda o nome como a pessoa digitou — com quebras de linha,
    espacos duplicados e acentuacao inconsistente. Exigir igualdade exata
    obrigaria a planilha a replicar esses detalhes invisiveis (ja aconteceu:
    um grupo so casou porque a celula tinha '\\n' no inicio).

    Normaliza: remove acentos, colapsa qualquer espaco/quebra em um unico
    espaco, e ignora maiusculas.
    """
    s = str(nome or "")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)          # \n, \t e espacos repetidos -> um espaco
    return s.strip().upper()


def _chave_nome_frouxa(nome):
    """Chave ainda mais tolerante: so letras e digitos.

    Usada como segunda tentativa — absorve diferencas de pontuacao,
    hifens e simbolos (ex.: '102 - MAX #SOC' vs '102 MAX SOC').
    """
    return re.sub(r"[^0-9A-Z]", "", _chave_nome(nome))


def indexar_por_nome(token=None, session=None, log=print):
    """Indice para casar com as colunas 'Contato Onvio' / 'Grupo Onvio'.

    Devolve {chave_normalizada: info}, com DUAS chaves por contato (normal e
    frouxa), para que resolver_por_nome tolere diferencas de digitacao.
    """
    indice = {}
    for c in listar_contatos(token=token, session=session, log=log):
        info = _info_contato(c)
        bruto = info["raw_name"] or ""
        chave = _chave_nome(bruto)
        if not chave:
            continue
        indice.setdefault(chave, info)
        frouxa = _chave_nome_frouxa(bruto)
        if frouxa:
            indice.setdefault("~" + frouxa, info)   # prefixo evita colisao
    return indice


def resolver_destinatario(codigo, indice):
    """Acha o contato de uma empresa pelo codigo. None se nao existir."""
    return indice.get(_norm_code(codigo))


def resolver_por_nome(nome, indice_nome):
    """Acha pelo nome exibido (como vem na planilha). None se nao existir.

    Duas tentativas: normalizada (sem acento, espacos colapsados) e frouxa
    (so letras e digitos). Assim a planilha nao precisa replicar quebras de
    linha ou espacos duplicados que existem no cadastro do Gestta.
    """
    if not nome or str(nome).strip().upper() == "NONE":
        return None
    chave = _chave_nome(nome)
    if not chave:
        return None
    achado = indice_nome.get(chave)
    if achado:
        return achado
    frouxa = _chave_nome_frouxa(nome)
    return indice_nome.get("~" + frouxa) if frouxa else None


def enviar_para(info, mensagem="", caminhos=None, desconsiderar=True,
                token=None, session=None, log=print):
    """Envia usando um `info` do indice — resolve sozinho contato vs grupo."""
    if not info or not info.get("id"):
        raise ValueError("Destinatario sem id (company_contact).")
    if info.get("is_group"):
        if not info.get("group_id"):
            raise ValueError("Grupo '%s' sem group_id: use a rota "
                             "/attendance-core/company/contact." % info.get("raw_name"))
        return enviar(info["id"], mensagem=mensagem, caminhos=caminhos,
                      group_id=info["group_id"], desconsiderar=desconsiderar,
                      token=token, session=session, log=log)
    return enviar(info["id"], mensagem=mensagem, caminhos=caminhos,
                  phone_number=info.get("telefone"), desconsiderar=desconsiderar,
                  token=token, session=session, log=log)


# ------------------------------------------------------------------ envio ----

def _destino(phone_number=None, group_id=None):
    """Monta a chave de destino do corpo: grupo tem precedencia sobre telefone.

    Contato -> {'phone_number': 5524999877658}  (inteiro, com 55, sem mascara)
    Grupo   -> {'group_id': '1203...@g.us'}
    """
    if group_id:
        return {"group_id": str(group_id)}
    if phone_number:
        digitos = re.sub(r"\D", "", str(phone_number))
        if not digitos:
            raise ValueError("Telefone sem digitos: %r" % (phone_number,))
        if not digitos.startswith("55"):
            digitos = "55" + digitos
        return {"phone_number": int(digitos)}
    raise ValueError("Informe phone_number (contato) ou group_id (grupo).")


def _sessao(token=None, session=None, log=print):
    token = token or obter_token("messenger", log=log)
    s = session or requests.Session()
    return token, s


def _checar(r, oque):
    """Levanta erro legivel; em 401 limpa o cache para forcar reautenticacao."""
    if r.status_code == 401:
        limpar_cache()
        raise RuntimeError("401 ao %s: token do Messenger invalido/expirado." % oque)
    if not r.ok:
        raise RuntimeError("Falha ao %s (HTTP %s): %s"
                           % (oque, r.status_code, (r.text or "")[:300]))
    return r


def enviar_texto(company_contact, mensagem, phone_number=None, group_id=None,
                 token=None, session=None, log=print):
    """POST /attendance-whatsapp/conversation/{contact}/text

    `company_contact` e o _id do contato no Messenger (vem do indice de contatos).
    Passe `phone_number` para contato individual OU `group_id` para grupo.
    Devolve o JSON da mensagem criada.
    """
    if requests is None:
        raise RuntimeError("A biblioteca 'requests' e necessaria (pip install requests).")
    if not (mensagem or "").strip():
        raise ValueError("Mensagem vazia. Para enviar so anexo, use enviar_arquivo.")

    token, s = _sessao(token, session, log)
    corpo = {"uuid": str(_uuid.uuid4()),
             "content": mensagem,
             "internal_message": False}
    corpo.update(_destino(phone_number, group_id))

    url = "%s/attendance-whatsapp/conversation/%s/text" % (API_BASE, company_contact)
    r = _checar(s.post(url, headers=_headers(token), json=corpo, timeout=TIMEOUT),
                "enviar texto")
    return r.json()


def upload_arquivo(company_contact, caminho, caption="",
                   token=None, session=None, log=print):
    """POST /attendance-documents/file/{contact}  (multipart)

    Primeiro dos dois passos do anexo. Devolve o JSON do arquivo, que alimenta
    enviar_arquivo(). Nao definir Content-Type: o requests monta o boundary.
    """
    if requests is None:
        raise RuntimeError("A biblioteca 'requests' e necessaria (pip install requests).")
    if not os.path.exists(caminho):
        raise FileNotFoundError("Arquivo nao encontrado: %s" % caminho)

    token, s = _sessao(token, session, log)
    tipo = mimetypes.guess_type(caminho)[0] or "application/octet-stream"
    url = "%s/attendance-documents/file/%s" % (API_BASE, company_contact)

    with open(caminho, "rb") as fh:
        arquivos = {"file": (os.path.basename(caminho), fh, tipo)}
        dados = {"company": COMPANY_ID, "caption": caption or ""}
        r = _checar(s.post(url,
                           headers={"authorization": token, "Accept": "application/json"},
                           files=arquivos, data=dados, timeout=TIMEOUT),
                    "subir arquivo")
    return r.json()


def enviar_arquivo(company_contact, caminho, caption="", phone_number=None,
                   group_id=None, token=None, session=None, log=print):
    """Envia um anexo: upload + POST .../file.

    `caption` e o texto que acompanha o arquivo — vazio envia so o anexo,
    que e o caso "arquivo sem mensagem" ja suportado pelo ONE.
    A resposta vem com status PENDING: o envio ao WhatsApp e assincrono.
    """
    token, s = _sessao(token, session, log)
    up = upload_arquivo(company_contact, caminho, caption=caption,
                        token=token, session=s, log=log)

    corpo = {"uuid": str(_uuid.uuid4()),
             "file_name": up.get("file_name"),
             "content_type": up.get("content_type"),
             "caption": up.get("caption", caption or ""),
             "file_url": up.get("file_url"),
             "_id": up.get("_id"),
             "internal_message": False,
             "mention_list": []}
    corpo.update(_destino(phone_number, group_id))

    url = "%s/attendance-whatsapp/conversation/%s/file" % (API_BASE, company_contact)
    r = _checar(s.post(url, headers=_headers(token), json=corpo, timeout=TIMEOUT),
                "enviar arquivo")
    return r.json()


def desconsiderar_atendimento(company_contact, token=None, session=None, log=print):
    """POST /attendance-core/attendance/disconsider

    Equivale ao 'desconsiderar' da UI, que no Selenium custa ~85 linhas de
    tratamento de bug de transferencia (ONE_V3.1.py:207-291).
    """
    if requests is None:
        raise RuntimeError("A biblioteca 'requests' e necessaria (pip install requests).")
    token, s = _sessao(token, session, log)
    url = "%s/attendance-core/attendance/disconsider" % API_BASE
    r = _checar(s.post(url, headers=_headers(token),
                       json={"company_contact": company_contact}, timeout=TIMEOUT),
                "desconsiderar atendimento")
    return r.json()


def enviar(company_contact, mensagem="", caminhos=None, phone_number=None,
           group_id=None, desconsiderar=True, token=None, session=None, log=print):
    """Envio completo de um destinatario — equivalente ao enviar_mensagem() do ONE.

    Cobre os casos que o ONE ja suporta:
      - so texto                     -> enviar_texto
      - texto + anexo(s)             -> 1o anexo leva o texto como caption
      - so anexo(s), sem mensagem    -> caption vazio
    Ao final, desconsidera o atendimento (como faz a UI).

    Devolve {'ok': bool, 'enviados': int, 'erros': [str]}.
    """
    token, s = _sessao(token, session, log)
    caminhos = [c for c in (caminhos or []) if c]
    texto = (mensagem or "").strip()
    resultado = {"ok": False, "enviados": 0, "erros": []}

    try:
        if not texto and not caminhos:
            raise ValueError("Nada a enviar: sem mensagem e sem arquivos.")

        if caminhos:
            # O texto vai como caption do primeiro anexo, espelhando a UI.
            for i, caminho in enumerate(caminhos):
                enviar_arquivo(company_contact, caminho,
                               caption=texto if i == 0 else "",
                               phone_number=phone_number, group_id=group_id,
                               token=token, session=s, log=log)
                resultado["enviados"] += 1
        else:
            enviar_texto(company_contact, texto,
                         phone_number=phone_number, group_id=group_id,
                         token=token, session=s, log=log)
            resultado["enviados"] += 1

        resultado["ok"] = True
    except Exception as e:
        resultado["erros"].append(str(e)[:300])
        log("Erro no envio: %s" % str(e)[:200])

    if desconsiderar:
        try:
            desconsiderar_atendimento(company_contact, token=token, session=s, log=log)
        except Exception as e:
            # Nao invalida o envio: a mensagem ja saiu.
            resultado["erros"].append("desconsiderar: %s" % str(e)[:200])
            log("Aviso: falha ao desconsiderar -> %s" % str(e)[:150])

    return resultado


# ------------------------------------------------------------- diagnostico ---

def testar_conexao(log=print):
    """Checa se o token funciona e mostra uma amostra dos contatos."""
    try:
        tok = obter_token("messenger", log=log)
    except RuntimeError as e:
        log("FALHA: %s" % e)
        return False

    horas = int((_expiracao(tok) - time.time()) / 3600)
    log("Token do Messenger obtido (~%dh de validade)." % horas)

    try:
        indice = indexar_por_codigo(token=tok, log=log)
    except Exception as e:
        log("FALHA ao listar contatos: %s" % str(e)[:200])
        return False

    if not indice:
        log("Conectou, mas nenhum contato foi indexado.")
        return False

    grupos = [c for c in indice.values() if c.get("is_group")]
    log("\nAmostra (5 primeiros codigos):")
    for cod in list(indice)[:5]:
        c = indice[cod]
        tipo = "GRUPO" if c.get("is_group") else "contato"
        alvo = c.get("group_id") or c.get("telefone") or "(sem destino)"
        log("  %-8s -> %-28s %-8s %s" % (cod, c["nome"][:28], tipo, alvo))
    log("\n%d codigos, dos quais %d sao grupos." % (len(indice), len(grupos)))
    if not grupos:
        log("Aviso: nenhum grupo indexado — envio para 'Grupo Onvio' nao vai funcionar.")
    return True


if __name__ == "__main__":
    testar_conexao()
