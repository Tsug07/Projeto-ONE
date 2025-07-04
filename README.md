
# 📬 AutoMessenger ONE

**AutoMessenger ONE** é uma aplicação desktop com interface gráfica desenvolvida em Python, projetada para **automatizar o envio de mensagens e anexos** para **contatos ou grupos no Onvio Messenger**, com base em dados estruturados em planilhas Excel. Ideal para departamentos de TI, RH, Financeiro ou Atendimento que buscam eficiência e padronização na comunicação via Onvio.

---

## 🧠 Funcionalidades

✅ Interface moderna e interativa com `customtkinter`  
✅ Automação de envio de mensagens e arquivos via navegador (Chrome + Selenium)  
✅ Suporte a múltiplos modelos de mensagens com estruturas específicas de Excel  
✅ Validação de dados, logs detalhados e barra de progresso visual  
✅ Sistema de mensagens customizáveis com edição via interface  
✅ Suporte a envio em massa com controle e personalização

---

## 🧰 Tecnologias Utilizadas

- [Python 3.10+](https://www.python.org/)
- [Selenium](https://selenium.dev/) – Automação de navegador
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) – Interface gráfica moderna
- [OpenPyXL](https://openpyxl.readthedocs.io/) – Leitura e validação de arquivos Excel
- [Pillow (PIL)](https://pillow.readthedocs.io/) – Manipulação de imagens
- [psutil](https://pypi.org/project/psutil/) – Gerenciamento de processos

---

## 📁 Modelos Suportados

| Modelo             | Campos Esperados no Excel                                                        | Tipo de Mensagem                          |
|--------------------|----------------------------------------------------------------------------------|-------------------------------------------|
| `ONE`              | Código, Empresa, Contato Onvio, Grupo Onvio, Caminho                             | Envio com anexos personalizados           |
| `ALL`              | Código, Empresa, Contato Onvio, Grupo Onvio                                      | Mensagem padrão                           |
| `ProrContrato`     | Código, Contato Onvio, Grupo Onvio, Nome, Vencimento                             | Prorrogação de contrato                   |
| `Cobranca`         | Código, Empresa, Contato Onvio, Grupo Onvio, Valor, Vencimento, Carta de Aviso   | Aviso de cobrança com diferentes versões  |
| `ComuniCertificado`| Código, Empresa, Contato Onvio, Grupo Onvio, CNPJ, Vencimento, Carta de Aviso    | Certificado digital vencendo              |

---

## ⚙️ Instalação

### 1. Clone o repositório (opcional)
```bash
git clone https://github.com/seuusuario/AutoMessenger-ONE.git
cd AutoMessenger-ONE
```

### 2. Instale as dependências

Instale com:

```bash
pip install -r requirements.txt
```

---

## ▶️ Executando o Script

```bash
python ONE.py
```

Certifique-se de ter o Google Chrome instalado e que o perfil `C:\PerfisChrome\automacao\Profile 1` exista (ou será criado automaticamente na primeira execução).

---


## 📂 Estrutura de Arquivos Esperada

```
AutoMessenger-ONE/
├── ONE.py
├── mensagens.json
├── logoOne.ico
├── logoOne.png
├── requirements.txt
├── README.md
└── AutoMessengerONE_Logs/   ← Gerado automaticamente
```

---

## 🧩 Recursos Adicionais

- As mensagens são carregadas a partir de `mensagens.json` e podem ser editadas diretamente pela interface.
- O sistema mantém um log de execução detalhado para rastrear ações e erros.
- Mensagens podem conter variáveis dinâmicas como `{nome}`, `{parcelas}`, `{cnpj_formatado}`, `{vencimentos}` etc.

---

## 💡 Dicas

- Execute o botão **"Iniciar Chrome de Automação"** antes do processamento para garantir login no Onvio Messenger.
- Sempre valide o Excel antes do envio.
- Utilize a edição de mensagens para adaptar os textos conforme o modelo.

---

## 👨‍💻 Desenvolvedor

**Hugo L. Almeida** – Equipe de TI  
🔧 Suporte técnico e melhorias: [hugogule@gmail.com]

---

## 📝 Licença

Este projeto é de uso interno. Consulte o time de TI para mais informações sobre distribuição e licença.
