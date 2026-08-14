# API do Gestta Messenger — envio de mensagens

**Conclusão: VIÁVEL, cobertura total.** Texto, anexo, grupo e o "desconsiderar"
têm endpoint próprio. Nada do fluxo do ONE fica preso à interface.

Descoberto em 13/08/2026 interceptando o tráfego real da tela do chat
(`capturar_api_messenger.py`), mesmo método que o Libby usou para a API de
tarefas. Todas as chamadas abaixo retornaram **200** no teste.

---

## Autenticação

Header `Authorization` com o **token do Messenger** (`user-jwt`, ~753 chars),
**não** o token core (`ngStorage-jwt`, ~1025 chars) usado em `/core/*`.

```
Authorization: JWT eyJ....
Content-Type:  application/json
Accept:        application/json, text/plain, */*
```

Base: `https://api.gestta.com.br`

---

## O identificador do destinatário

Todas as rotas de conversa usam o **`company_contact`** — o `_id` do contato no
Messenger, o mesmo `_id` devolvido por
`GET /messenger-admin/company/contact` (já usado em `gestta_api.listar_contatos`).

É o mesmo id para contato e para grupo: o que muda é o corpo (`phone_number` vs
`group_id`), não a URL.

---

## 1) Enviar texto

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

- `uuid`: v4 gerado pelo cliente (idempotência/correlação). Um por mensagem.
- `phone_number`: **número inteiro**, com 55, sem máscara.
- `group_id`: string terminada em `@g.us`.
- Resposta: `{_id, external_id, type:"OUTBOUND", content_type:"CHAT", status:"SENT", ...}`

---

## 2) Enviar arquivo — dois passos

**2.1 Upload** (multipart, sem `Content-Type` manual — o requests monta o boundary):
```
POST /attendance-documents/file/{company_contact}     -> 200
```
Campos: `file` (o arquivo), `company` (id da empresa), `caption` (texto que
acompanha o anexo).

Resposta:
```json
{"_id":"...","file_name":"logoCanella.png","file_url":"https://api.gestta.com.br/attendance-documents/file/{contact}/{fileId}","content_type":"image/png","caption":"teste\r\n","size":49244}
```

**2.2 Enviar** (JSON, repassando o que o upload devolveu):
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

Resposta: `content_type:"LINK"`, `status:"PENDING"` (o envio ao WhatsApp é
assíncrono — PENDING aqui é normal, não é erro).

> O `caption` é o texto da mensagem que acompanha o anexo. Isso cobre o caso do
> ONE de "enviar arquivo sem mensagem" (caption vazio) e "com mensagem".

---

## 3) Desconsiderar atendimento

O passo que na UI custa ~85 linhas de tratamento de bug (`ONE_V3.1.py:207-291`)
é **uma chamada**:

```
POST /attendance-core/attendance/disconsider     -> 200
{"company_contact": "67b71fa7dca30e3d2a697214"}
```

---

## 4) Auxiliares (opcionais)

Vistos na captura, disparados pela UI. Não são necessários para enviar:
```
PATCH /attendance-core/notification/{company_contact}/seen    {}
PATCH /attendance-core/conversation/{company_contact}/seen    {}
```

---

## Company id

`company: "6418e03fe478c70006b05005"` — constante da Canella & Santos, exigida
no multipart do upload. Vem também em toda resposta de envio.

---

## O que isso elimina do ONE

| Hoje (Selenium) | Via API |
|---|---|
| Buscar contato por nome na tela, aba contato vs grupo | `company_contact` resolvido pelo índice de contatos |
| Inserir HTML na barra e clicar em enviar | `POST .../text` |
| `input[type=file]` + clique no botão | upload + `POST .../file` |
| Desconsiderar + tratamento de bug de transferência | `POST .../disconsider` |
| Chrome aberto, perfil, keep-alive, sessão expirada | token com renovação |

---

## Casamento de nome (planilha -> Messenger)

O Gestta guarda o nome **como a pessoa digitou**: com espaços nas pontas,
espaços duplicados e até quebras de linha. Um grupo real está cadastrado como
`' 102 - TESTE T.I CANELLA '`.

Por isso `resolver_por_nome` normaliza antes de comparar, em duas tentativas:

1. **normalizada** — sem acentos, espaços/quebras colapsados, maiúsculas
2. **frouxa** — só letras e dígitos (absorve `-`, `.`, `#`)

Assim a planilha não precisa replicar caracteres invisíveis do cadastro.
Validado contra a base real: 2258 nomes indexados.

O que ainda **não** resolve: nome genuinamente diferente (apelido, pessoa
trocada, empresa renomeada). Esses aparecem como SEM DESTINO no dry-run e
exigem correção no cadastro ou na planilha.

---

## Cuidados

- **Sem UI não há freio.** O `DELAY_ENTRE_ENVIOS` passa a ser o único
  espaçamento; via API o loop roda muito mais rápido que pela tela.
- `status:"PENDING"` no anexo é assíncrono — não confundir com falha.
- Um `uuid` novo por mensagem; reaproveitar pode deduplicar no servidor.
- Token expira (~24h). Em 401, limpar cache e reautenticar.
