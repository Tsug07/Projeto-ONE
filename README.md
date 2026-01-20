# AutoMessenger ONE

<div align="center">

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-Automation-43B02A?style=for-the-badge&logo=selenium&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-GUI-blue?style=for-the-badge)
![License](https://img.shields.io/badge/License-Internal_Use-red?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Active-success?style=for-the-badge)
![Version](https://img.shields.io/badge/Version-3.0-orange?style=for-the-badge)

<br>

**Solução de automação corporativa para envio de mensagens e anexos via Onvio Messenger**

[Funcionalidades](#-funcionalidades) •
[Instalação](#%EF%B8%8F-instalação) •
[Uso](#-uso) •
[Modelos](#-modelos-suportados) •
[Configuração](#-configuração)

</div>

---

## Sobre o Projeto

**AutoMessenger ONE** é uma aplicação desktop desenvolvida em Python com interface gráfica moderna, projetada para automatizar o envio de mensagens e anexos para contatos ou grupos no **Onvio Messenger**. A ferramenta utiliza dados estruturados em planilhas Excel, sendo ideal para departamentos de TI, RH, Financeiro ou Atendimento que buscam eficiência e padronização na comunicação corporativa.

---

## Funcionalidades

| Recurso | Descrição |
|---------|-----------|
| **Interface Moderna** | GUI desenvolvida com CustomTkinter, suporte a tema Dark/Light |
| **Automação Robusta** | Integração com Chrome via Selenium WebDriver |
| **Múltiplos Modelos** | Suporte a diferentes estruturas de mensagens e Excel |
| **Agendamento** | Envio programado com contagem regressiva visual |
| **Validação de Dados** | Verificação automática de planilhas antes do envio |
| **Logs Detalhados** | Sistema completo de logging para auditoria |
| **Editor de Mensagens** | Personalização de templates diretamente na interface |
| **Envio de Anexos** | Suporte a múltiplos arquivos por mensagem |
| **Keep-Alive** | Manutenção automática de sessão do navegador |
| **Multi-Perfil** | Suporte a múltiplos perfis do Chrome |

---

## Tecnologias Utilizadas

<div align="center">

| Tecnologia | Versão | Propósito |
|------------|--------|-----------|
| [Python](https://www.python.org/) | 3.10+ | Linguagem principal |
| [Selenium](https://selenium.dev/) | Latest | Automação de navegador |
| [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) | Latest | Interface gráfica |
| [OpenPyXL](https://openpyxl.readthedocs.io/) | Latest | Manipulação de Excel |
| [Pillow](https://pillow.readthedocs.io/) | Latest | Processamento de imagens |
| [psutil](https://pypi.org/project/psutil/) | Latest | Gerenciamento de processos |
| [webdriver-manager](https://pypi.org/project/webdriver-manager/) | Latest | Gerenciamento do ChromeDriver |

</div>

---

## Modelos Suportados

| Modelo | Campos do Excel | Caso de Uso |
|--------|-----------------|-------------|
| `ONE` | Código, Empresa, Contato Onvio, Grupo Onvio, Caminho | Envio com anexos personalizados |
| `ALL` | Codigo, Empresa, Contato Onvio, Grupo Onvio | Mensagem padrão em massa |
| `ALL_info` | Codigo, Empresa, Contato Onvio, Grupo Onvio, Competencia | Mensagem com competência |
| `Cobranca` | Código, Empresa, Contato Onvio, Grupo Onvio, Valor, Vencimento, Carta de Aviso | Avisos de cobrança |
| `ComuniCertificado` | Codigo, Empresa, Contato Onvio, Grupo Onvio, CNPJ, Vencimento, Carta de Aviso | Certificado digital |

---

## Instalação

### Pré-requisitos

- Python 3.10 ou superior
- Google Chrome instalado
- Acesso ao Onvio Messenger

### Passo a Passo

```bash
# 1. Clone o repositório
git clone https://github.com/seuusuario/AutoMessenger-ONE.git
cd AutoMessenger-ONE

# 2. Crie um ambiente virtual (recomendado)
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac

# 3. Instale as dependências
pip install -r requirements.txt
```

---

## Uso

### Iniciando a Aplicação

```bash
python ONE_V3.py
```

### Fluxo de Trabalho

```
1. Iniciar Chrome de Automação → Login no Onvio Messenger
         ↓
2. Selecionar Modelo → Escolher tipo de mensagem
         ↓
3. Carregar Planilha Excel → Validação automática
         ↓
4. Configurar Opções → Anexos, agendamento (opcional)
         ↓
5. Processar Envio → Monitorar progresso via logs
```

### Dicas Importantes

> **Primeiro uso:** Execute o botão "Iniciar Chrome de Automação" e faça login no Onvio Messenger antes de processar.

> **Validação:** Sempre valide a planilha Excel antes do envio em massa.

> **Anexos (Modelo ONE):** Os arquivos devem estar em `Documentos\Relatorios`.

---

## Configuração

### Estrutura do Projeto

```
AutoMessenger-ONE/
├── ONE_V3.py              # Aplicação principal
├── mensagens.json         # Templates de mensagens
├── logoOne.ico            # Ícone da aplicação
├── logoOne.png            # Logo da interface
├── requirements.txt       # Dependências Python
├── README.md              # Documentação
└── AutoMessengerONE_Logs/ # Logs de execução (gerado automaticamente)
```

### Perfis do Chrome

Os perfis de automação são armazenados em:
```
C:\PerfisChrome\automacao\Profile 1
C:\PerfisChrome\automacao\Profile 2
```

### Variáveis Dinâmicas

As mensagens suportam variáveis que são substituídas automaticamente:

| Variável | Descrição |
|----------|-----------|
| `{nome}` | Nome do contato/empresa |
| `{empresa}` | Nome da empresa |
| `{valor}` | Valor da parcela |
| `{vencimento}` | Data de vencimento |
| `{cnpj_formatado}` | CNPJ formatado |
| `{competencia}` | Mês/ano de competência |

---

## Logs e Monitoramento

Os logs são salvos automaticamente em `AutoMessengerONE_Logs/` com o formato:
```
log_YYYYMMDD_HHMMSS.txt
```

Cada log contém:
- Timestamp de cada ação
- Status de envio por destinatário
- Erros e exceções detalhados
- Resumo final de processamento

---

## Solução de Problemas

| Problema | Solução |
|----------|---------|
| Chrome não inicia | Verifique se o Chrome está instalado e atualizado |
| Erro de perfil | Delete a pasta do perfil e reinicie a aplicação |
| Anexo não encontrado | Verifique se o arquivo existe em `Documentos\Relatorios` |
| Timeout na página | Aumente o tempo de espera ou verifique a conexão |
| Sessão expirada | Use o recurso Keep-Alive ou reinicie o Chrome |

---

## Desenvolvedor

<div align="center">

**Hugo L. Almeida**
Equipe de TI

[![Email](https://img.shields.io/badge/Email-hugoalmeida.canellaesantos%40gmail.com-D14836?style=flat-square&logo=gmail&logoColor=white)](mailto:hugoalmeida.canellaesantos@gmail.com)

</div>

---

## Licença

<div align="center">

Este projeto é de **uso interno corporativo**.
Consulte o time de TI para informações sobre distribuição.

---

*Desenvolvido com Python para automação corporativa*

</div>
