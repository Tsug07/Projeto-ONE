# API do Gestta Messenger — envio de mensagens (WhatsApp)

**Conclusão: VIÁVEL, cobertura total.** Texto, anexo, grupo e o "desconsiderar"
têm endpoint próprio. Nada do fluxo depende da interface.

Investigado ao vivo em **13/08/2026**, interceptando o tráfego real da tela do
chat (hook de fetch/XHR na página logada) — mesmo método usado em
`API_GESTTA_TAREFAS.md`. Todas as chamadas abaixo retornaram **200**.

**Validado em produção:** 22 mensagens de cobrança enviadas em 2min24s
(≈6,5s por mensagem, incluindo o delay configurado). Contato e grupo.

---

## Autenticação — atenção ao token

Aqui está a pegadinha que custa 401: **são dois tokens diferentes**, e o do
Messenger **não** é o mesmo do `/core/*` já usado nas tarefas.

| Uso | localStorage | Tamanho | Rotas |
|---|---|---|---|
| Core / tarefas | `ngStorage-jwt` | ~1025 chars | `/core/*` |
| **Messenger / chat** | **`user-jwt`** | ~753 chars | `/attendance-*`, `/messenger-admin/*` |

O Libby já distingue os dois: `gestta_auth.obter_token_messenger()` lê o
`MESSENGER_TOKEN`, populado por `onvio_session.preparar_tokens_env()` e
renovado por `token_service`. **Reaproveitar direto** — nada novo é necessário
do lado da autenticação.

```
Authorization: JWT eyJ....        (valor do user-jwt, verbatim)
Content-Type:  application/json
Accept:        application/json, text/plain, */*
```

Base: `https://api.gestta.com.br`

---

## O identificador do destinatário

Todas as rotas de conversa usam o **`company_contact`** — o `_id` do contato no
Messenger.

**É o mesmo id para contato e para grupo.** O que muda é o corpo da requisição:
`phone_number` para contato individual, `group_id` para grupo.

---

## 1) Listar contatos e grupos

```
GET /attendance-core/company/contact        -> array (sem paginação)
```

Resposta — contato:
```json
{"_id":"6a7dcc33e972ae16306b7c8e","name":"102 - Hugo","phone_number":5524999877658,
 "company":"6418e03fe478c70006b05005"}
```

Resposta — grupo:
```json
{"_id":"67b71fa7dca30e3d2a697214","name":" 102 - TESTE T.I CANELLA ",
 "is_group":true,"group_id":"120363404926085233@g.us"}
```

> **Importante para o Libby.** O `contatos/messenger_onvio.py` usa hoje
> `GET /messenger-admin/company/contact` (paginado, `{docs, hasNextPage}`).
> Essa rota **não devolve `is_group` nem `group_id`** — ou seja, com ela é
> impossível enviar para grupo. A rota `/attendance-core/*` acima devolve
> contatos **e** grupos no mesmo array, com os dois campos.
>
> Para o `atualizar_contatos` passar a cobrir grupos, migrar para essa rota e
> gravar o `group_id` junto.

Base real (13/08/2026): **2346 registros, sendo 258 grupos.**

Auxiliar, se precisar dos membros de um grupo:
```
GET /attendance-whatsapp/groups/{company_contact}/participants
```

---

## 2) Enviar texto

```
POST /attendance-whatsapp/conversation/{company_contact}/text     -> 200
```

Contato individual:
```json
{
  "uuid": "5386ab04-23b0-4141-afc4-8d210505c17f",
  "content": "teste",
  "phone_number": 5524999877658,
  "internal_message": false
}
```

Grupo — troca `phone_number` por `group_id`:
```json
{
  "uuid": "79df477a-4556-44e8-a889-23da5dfe4070",
  "content": "teste",
  "group_id": "120363404926085233@g.us",
  "internal_message": false
}
```

- `uuid`: v4 gerado pelo cliente, **um novo por mensagem** (idempotência).
- `phone_number`: **inteiro**, com 55, sem máscara.
- `group_id`: string terminada em `@g.us`.
- `content`: texto puro; `\n` funciona normalmente (nada de HTML).

Resposta: `{_id, external_id, type:"OUTBOUND", content_type:"CHAT", status:"SENT", ...}`

---

## 3) Enviar arquivo — dois passos

**3.1 Upload** (multipart; não definir `Content-Type` manualmente):
```
POST /attendance-documents/file/{company_contact}     -> 200
```
Campos: `file` (binário), `company` (id da empresa), `caption` (texto que
acompanha o anexo).

Resposta:
```json
{"_id":"6a7dcff52fcfeef1db0cf864","file_name":"logoCanella.png",
 "file_url":"https://api.gestta.com.br/attendance-documents/file/{contact}/{fileId}",
 "content_type":"image/png","caption":"teste\r\n","size":49244}
```

**3.2 Enviar** (JSON, repassando o que o upload devolveu):
```
POST /attendance-whatsapp/conversation/{company_contact}/file     -> 200
```
```json
{
  "uuid": "67b97087-d31c-4885-8082-8884f39e8ed3",
  "file_name": "logoCanella.png",
  "content_type": "image/png",
  "caption": "teste\r\n",
  "file_url": "https://api.gestta.com.br/attendance-documents/file/{contact}/{fileId}",
  "_id": "{fileId}",
  "internal_message": false,
  "mention_list": []
}
```
Para grupo, acrescentar `group_id` (mesma regra do texto).

Resposta: `content_type:"LINK"`, `status:"PENDING"` — o envio ao WhatsApp é
assíncrono. **PENDING aqui é normal, não é falha.**

> O `caption` é o texto que acompanha o arquivo. Caption vazio envia só o
> anexo; com texto, envia anexo + mensagem numa tacada.

---

## 4) Desconsiderar atendimento

Enviar uma mensagem **abre um atendimento** no Gestta. A UI encerra logo em
seguida, e é isso que o "desconsiderar" faz:

```
POST /attendance-core/attendance/disconsider     -> 200
{"company_contact": "67b71fa7dca30e3d2a697214"}
```

Sem isso, os atendimentos ficam abertos na fila da equipe. **Chamar sempre
após o envio.**

(No ONE via Selenium esse passo custava ~85 linhas de tratamento de bug de
transferência. Via API é uma chamada.)

---

## 5) Auxiliares (opcionais)

Disparados pela UI, não necessários para enviar:
```
PATCH /attendance-core/notification/{company_contact}/seen    {}
PATCH /attendance-core/conversation/{company_contact}/seen     {}
GET   /attendance-core/conversation/{company_contact}/message?limit=30&page=1
GET   /attendance-core/conversation/all
```

---

## Company id

`company: "6418e03fe478c70006b05005"` — Canella & Santos. Exigido no multipart
do upload; vem também em toda resposta de envio.

---

## Casamento de nome (planilha → Messenger)

O Gestta guarda o nome **como a pessoa digitou**: espaços nas pontas, espaços
duplicados, às vezes quebra de linha. Um grupo real está cadastrado como
`' 102 - TESTE T.I CANELLA '`.

Casar por igualdade exata **falha na prática** — obriga a planilha a replicar
caracteres invisíveis. Normalizar antes de comparar, em duas tentativas:

1. **normalizada** — sem acentos, `\s+` colapsado em um espaço, maiúsculas
2. **frouxa** — só letras e dígitos (absorve `-`, `.`, `#`)

Implementação de referência: `gestta_api._chave_nome` / `resolver_por_nome`
(projeto ONE). Validado contra a base real: 2258 nomes indexados.

**O Libby resolve isso pela raiz.** O `messenger_onvio.py` indexa por *código
de empresa* extraído do próprio nome do contato (`09/55/884 - JOSIAS` →
`['9','55','884']`). Planilha gerada pelo Libby já nasce consistente com o
cadastro — não precisa de tolerância.

Ressalva do índice por código: um contato costuma responder por várias
empresas (ex.: `JANICE #SOC #AMAR` cobre 5 códigos), então códigos distintos
apontam para a mesma pessoa. Correto, mas conte destinatários, não linhas.

---

## Ordem de resolução do destinatário

Espelhando a regra que o ONE já usava na interface:

1. `Grupo Onvio` preenchido e ≠ `NONE` → grupo, por nome
2. senão `Contato Onvio` ≠ `NONE` → contato, por nome
3. senão código da empresa (fallback)

---

## Fluxo completo de um envio

```
1. token   = MESSENGER_TOKEN (user-jwt)          [token_service do Libby]
2. índice  = GET /attendance-core/company/contact  (uma vez por rodada)
3. destino = resolve nome/código -> {id, is_group, group_id, phone_number}
4. envio:
     texto  -> POST /attendance-whatsapp/conversation/{id}/text
     anexo  -> POST /attendance-documents/file/{id}          (multipart)
               POST /attendance-whatsapp/conversation/{id}/file
5. POST /attendance-core/attendance/disconsider  {"company_contact": id}
6. pausa entre destinatários
```

---

## Cuidados

- **Sem UI não há freio.** Pela tela, cada envio carregava ~5s de esperas
  implícitas. Via API o loop dispara muito mais rápido — o delay entre
  destinatários passa a ser o **único** espaçamento. Medido: 22 mensagens em
  2min24s com delay de 3s.
- **`disconsider` sempre**, senão os atendimentos ficam abertos na fila.
- **`uuid` novo por mensagem** — reaproveitar pode deduplicar no servidor.
- **`status:"PENDING"`** no anexo é assíncrono, não é erro.
- **Token expira (~24h).** Em 401, limpar cache e reautenticar (o
  `token_service` do Libby já faz isso).
- **Envios são reais e irreversíveis.** Vale um modo dry-run que resolve
  destinatários e monta as mensagens sem enviar — foi o que pegou os
  problemas de casamento de nome antes de qualquer disparo.

---

## Referências

- `gestta_api.py` (projeto ONE) — implementação de referência: auth, índice de
  contatos, `enviar_texto`, `enviar_arquivo`, `desconsiderar_atendimento`,
  `enviar_para`.
- `capturar_api_messenger.py` (projeto ONE) — hook de fetch/XHR usado nesta
  investigação; serve para mapear qualquer outra tela do Gestta.
- `API_GESTTA_TAREFAS.md` (Libby) — mesma abordagem, para a API de tarefas.
