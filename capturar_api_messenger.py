# -*- coding: utf-8 -*-
"""
capturar_api_messenger.py
-------------------------
Descobre o contrato da API de ENVIO de mensagem do Gestta Messenger.

Nenhum projeto nosso envia mensagem via API ainda: o Libby so LE contatos
(messenger_onvio.py) e o runner de avisos so mexe em tarefas. Este script
fecha essa lacuna capturando a chamada real que a propria tela do Gestta faz.

Mesmo metodo que o Libby usou para descobrir a API de tarefas: instala um hook
leve em fetch/XHR na pagina logada, voce envia UMA mensagem de teste pela
interface, e o hook registra metodo, URL, headers e corpo.

COMO USAR
---------
  1. FECHE o Chrome de automacao do ONE (este script abre o seu proprio,
     usando o MESMO perfil, e o Chrome nao aceita dois processos no mesmo
     diretorio de perfil).
  2. python capturar_api_messenger.py           (perfil 1, padrao)
     python capturar_api_messenger.py --perfil 2
  3. O script abre o Gestta ja logado (a sessao vem do perfil). Quando
     aparecer "HOOK ATIVO", envie UMA mensagem de teste para um contato seu
     (nao para cliente!). Se quiser mapear anexo, envie tambem uma com
     arquivo; para mapear grupo, envie uma num grupo.
  4. Volte aqui e tecle ENTER. O resultado sai na tela e em
     captura_api_messenger.json.

SEGURANCA
---------
  - Envie so para um contato interno de teste. A mensagem e real.
  - O script NAO envia nada sozinho; ele so observa.
  - Tokens sao mascarados no arquivo de saida (nunca gravamos o JWT inteiro).
"""

import argparse
import json
import os
import sys
import time
from pathlib import Path

try:
    from selenium import webdriver
except ImportError:
    print("Selenium nao instalado. Rode: pip install selenium")
    sys.exit(1)

# Mesmo perfil que o ONE usa (ver obter_user_data_dir no ONE_V3.1.py), para
# herdar a sessao ja logada no Gestta.
PERFIL_DIR = r"C:\PerfisChrome\automacao_perfil%s"
URL_CHAT = "https://app.gestta.com.br/attendance/#/chat/contact-list"
SAIDA = Path(__file__).parent / "captura_api_messenger.json"

# So registra chamadas que interessam (evita ruido de telemetria/assets).
FILTRO = "gestta.com.br"

# Hook: envelopa window.fetch e XMLHttpRequest, guardando em window.__ONE_CAPTURA.
# Mascara qualquer header de autorizacao antes de guardar.
JS_HOOK = r"""
(function () {
  if (window.__ONE_CAPTURA) { return 'ja-instalado'; }
  window.__ONE_CAPTURA = [];
  var FILTRO = %s;
  var MAX = 400;

  function mascarar(v) {
    if (!v) return v;
    var s = String(v);
    if (s.length > 24 && /^(JWT|Bearer)?\s*ey/i.test(s)) {
      return s.slice(0, 12) + '...[MASCARADO ' + s.length + ' chars]';
    }
    return s;
  }

  function guardar(reg) {
    try {
      if (!reg.url || reg.url.indexOf(FILTRO) === -1) return;
      if (window.__ONE_CAPTURA.length >= MAX) return;
      window.__ONE_CAPTURA.push(reg);
    } catch (e) {}
  }

  function corpo(b) {
    try {
      if (b == null) return null;
      if (typeof b === 'string') return b.slice(0, 4000);
      if (b instanceof FormData) {
        var campos = [];
        b.forEach(function (v, k) {
          campos.push(k + '=' + (v instanceof File
            ? '[FILE nome=' + v.name + ' tipo=' + v.type + ' bytes=' + v.size + ']'
            : String(v).slice(0, 200)));
        });
        return 'FormData{' + campos.join('&') + '}';
      }
      if (b instanceof Blob) return '[Blob ' + b.size + ' bytes]';
      return JSON.stringify(b).slice(0, 4000);
    } catch (e) { return '[corpo nao serializavel]'; }
  }

  // ---- fetch ----
  var fetchOrig = window.fetch;
  window.fetch = function (entrada, init) {
    var url = (typeof entrada === 'string') ? entrada : (entrada && entrada.url) || '';
    var metodo = (init && init.method) || (entrada && entrada.method) || 'GET';
    var headers = {};
    try {
      var h = (init && init.headers) || (entrada && entrada.headers);
      if (h) {
        if (typeof h.forEach === 'function') h.forEach(function (v, k) { headers[k] = mascarar(v); });
        else Object.keys(h).forEach(function (k) { headers[k] = mascarar(h[k]); });
      }
    } catch (e) {}
    var reg = { via: 'fetch', metodo: metodo, url: url, headers: headers,
                corpo: corpo(init && init.body), quando: new Date().toISOString() };
    guardar(reg);
    return fetchOrig.apply(this, arguments).then(function (resp) {
      try {
        reg.status = resp.status;
        resp.clone().text().then(function (t) { reg.resposta = (t || '').slice(0, 2000); }).catch(function () {});
      } catch (e) {}
      return resp;
    });
  };

  // ---- XMLHttpRequest ----
  var abrirOrig = XMLHttpRequest.prototype.open;
  var enviarOrig = XMLHttpRequest.prototype.send;
  var setHdrOrig = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function (metodo, url) {
    this.__one = { via: 'xhr', metodo: metodo, url: url, headers: {} };
    return abrirOrig.apply(this, arguments);
  };
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
    try { if (this.__one) this.__one.headers[k] = mascarar(v); } catch (e) {}
    return setHdrOrig.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function (body) {
    var self = this;
    if (self.__one) {
      self.__one.corpo = corpo(body);
      self.__one.quando = new Date().toISOString();
      guardar(self.__one);
      self.addEventListener('load', function () {
        try {
          self.__one.status = self.status;
          self.__one.resposta = String(self.responseText || '').slice(0, 2000);
        } catch (e) {}
      });
    }
    return enviarOrig.apply(this, arguments);
  };

  return 'instalado';
})();
""" % json.dumps(FILTRO)

JS_COLETA = "return JSON.stringify(window.__ONE_CAPTURA || []);"


def abrir_chrome(perfil):
    """Abre o Chrome no MESMO perfil do ONE, para herdar a sessao logada.

    O ONE nao abre o Chrome com --remote-debugging-port, entao nao ha porta
    para anexar; abrimos nosso proprio processo apontando para o perfil.
    """
    user_data_dir = PERFIL_DIR % perfil
    if not os.path.isdir(user_data_dir):
        print("Perfil nao encontrado: %s" % user_data_dir)
        print("Rode o ONE e abra o Chrome de automacao ao menos uma vez,")
        print("ou use --perfil com o numero correto.")
        sys.exit(1)

    print("Usando perfil %s (%s)" % (perfil, user_data_dir))
    opts = webdriver.ChromeOptions()
    opts.add_argument("--user-data-dir=%s" % user_data_dir)
    opts.add_argument("--start-maximized")
    opts.add_argument("--lang=pt-BR")
    opts.add_argument("--disable-translate")
    try:
        return webdriver.Chrome(options=opts)
    except Exception as e:
        msg = str(e)
        print("\nNao consegui abrir o Chrome com esse perfil.")
        if "user data directory is already in use" in msg.lower():
            print("O perfil ja esta em uso: FECHE o Chrome de automacao do ONE")
            print("(e o proprio ONE, se estiver com o Chrome aberto) e tente de novo.")
        else:
            print("Detalhe: %s" % msg[:250])
        sys.exit(1)


def ir_para_gestta(drv):
    """Deixa o driver numa aba do Gestta; abre a tela do chat se preciso."""
    for h in drv.window_handles:
        drv.switch_to.window(h)
        if "gestta.com.br" in (drv.current_url or ""):
            return True
    drv.get(URL_CHAT)
    print("Abrindo o Gestta... aguarde o carregamento.")
    for _ in range(30):
        time.sleep(1)
        if "gestta.com.br" in (drv.current_url or ""):
            time.sleep(3)   # deixa o app terminar de subir
            return True
    return False


def instalar_hook(drv):
    return drv.execute_script(JS_HOOK)


def coletar(drv):
    bruto = drv.execute_script(JS_COLETA)
    try:
        return json.loads(bruto or "[]")
    except Exception:
        return []


def interessante(reg):
    """Chamadas de escrita (POST/PUT/PATCH) sao as candidatas a envio."""
    return (reg.get("metodo") or "").upper() in ("POST", "PUT", "PATCH")


def resumir(regs):
    print("\n" + "=" * 70)
    print("CAPTURA: %d chamadas ao Gestta" % len(regs))
    print("=" * 70)

    escritas = [r for r in regs if interessante(r)]
    if not escritas:
        print("\nNenhuma chamada de escrita capturada.")
        print("A mensagem chegou a ser enviada com o hook ativo?")
        return

    print("\n%d chamadas de escrita (candidatas ao envio):\n" % len(escritas))
    for i, r in enumerate(escritas, 1):
        print("-" * 70)
        print("[%d] %s %s" % (i, r.get("metodo"), r.get("url")))
        if r.get("status"):
            print("     status: %s" % r["status"])
        if r.get("corpo"):
            print("     corpo: %s" % str(r["corpo"])[:600])
        if r.get("resposta"):
            print("     resposta: %s" % str(r["resposta"])[:300])


def main():
    ap = argparse.ArgumentParser(description="Captura a API de envio do Gestta Messenger.")
    ap.add_argument("--perfil", default="1",
                    help="Perfil do Chrome do ONE (1 ou 2). Padrao: 1")
    args = ap.parse_args()

    print(__doc__)
    drv = abrir_chrome(args.perfil)
    try:
        if not ir_para_gestta(drv):
            print("Nao consegui abrir o Gestta. A sessao deste perfil esta logada?")
            sys.exit(1)

        print("Gestta aberto: %s" % (drv.current_url or "")[:90])
        estado = instalar_hook(drv)
        print("Hook: %s" % estado)
        print("\n" + "=" * 70)
        print("HOOK ATIVO")
        print("=" * 70)
        print("Agora, NO CHROME que acabou de abrir, envie mensagens de teste:")
        print("  1. Uma para um CONTATO interno (nao cliente!)  -> envio de texto")
        print("  2. Uma com ARQUIVO anexado                      -> upload")
        print("  3. Uma para um GRUPO                            -> grupos tem API?")
        print("\nATENCAO: nao recarregue a pagina - o hook se perde no refresh.")
        print("=" * 70)
        input("\nDepois de enviar, tecle ENTER para coletar... ")

        regs = coletar(drv)
        resumir(regs)

        SAIDA.write_text(json.dumps(regs, ensure_ascii=False, indent=2), encoding="utf-8")
        print("\n\nCaptura completa salva em: %s" % SAIDA)
        print("Me mande esse arquivo (ou o trecho da chamada de envio) para")
        print("montarmos a camada de API do ONE.")
    finally:
        print("\nO Chrome fica aberto para voce conferir; feche quando quiser.")


if __name__ == "__main__":
    main()
