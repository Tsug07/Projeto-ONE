# Projeto-ONE
Automessenger ONE
Versão 1.0 | Desenvolvido por Hugo L. Almeida - Equipe de TI
Automessenger ONE é uma ferramenta de automação unificada para envio de mensagens via Onvio Messenger. Ele suporta múltiplos modelos de mensagens, permitindo personalização através de arquivos Excel e templates de mensagens definidos em mensagens.json. A interface gráfica, construída com CustomTkinter, facilita a interação com o usuário, incluindo seleção de modelos, arquivos Excel, edição de mensagens e monitoramento de progresso.
Funcionalidades Principais

Seleção de Modelo: Escolha entre modelos como Prorrogação Contrato, Cobrança (1 a 6) e Certificado (1 a 4).
Seleção de Excel: Carregue arquivos Excel com dados estruturados para cada modelo.
Escolha da Linha Inicial: Defina a linha inicial do Excel para processamento.
Seleção de Mensagem: Escolha mensagens predefinidas de mensagens.json.
Adicionar/Editar Mensagem: Crie ou modifique templates de mensagens.
Remover Mensagem: Exclua mensagens desnecessárias.
Iniciar Chrome de Automação: Abra o Chrome para automação de envio de mensagens.
Iniciar Processamento: Processe os dados do Excel e envie mensagens automaticamente.
Cancelar Processamento: Interrompa o processo de envio, se necessário.
Fechar Programa: Saia do aplicativo com segurança.
Abrir Log: Visualize o arquivo de log para auditoria.
Registro de Log: Todas as ações e erros são registrados automaticamente.

Requisitos

Sistema Operacional: Windows (devido ao uso de caminhos específicos como C:\PerfisChrome\automacao).
Python: Versão 3.7 ou superior.
Dependências:
customtkinter
selenium
webdriver_manager
openpyxl
psutil
Navegador Chrome instalado.


Arquivos Necessários:
mensagens.json: Contém os templates de mensagens.
Arquivo Excel formatado corretamente para o modelo selecionado.



Instalação

Clone o Repositório (se aplicável):
git clone <URL_DO_REPOSITORIO>
cd automessenger-one


Crie um Ambiente Virtual (opcional, mas recomendado):
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate  # Windows


Instale as Dependências:
pip install customtkinter selenium webdriver-manager openpyxl psutil


Configure o Arquivo mensagens.json:

Certifique-se de que o arquivo mensagens.json está no mesmo diretório do script principal (ONE.py).
Verifique se ele contém os templates de mensagens necessários (exemplo fornecido abaixo).


Configure o Chrome:

Certifique-se de que o Chrome está instalado.
O script usa um perfil específico do Chrome em C:\PerfisChrome\automacao\Profile 1. Faça login no Onvio Messenger antes de iniciar o processamento.



Modelos Suportados
1. Prorrogação Contrato

Objetivo: Notifica clientes sobre contratos de funcionários que estão prestes a vencer.
Colunas Obrigatórias no Excel:
Codigo: Código da empresa.
Contato Onvio: Nome do contato no Onvio.
Grupo Onvio: Nome do grupo no Onvio.
Nome: Nome da empresa.
Vencimento: Data de vencimento do contrato.


Mensagem: Usa o template "Prorrogação Contrato" de mensagens.json, formatando a lista de funcionários e vencimentos.

2. Cobrança

Objetivo: Envia lembretes de pagamento para honorários contábeis atrasados (6 níveis de escalonamento).
Colunas Obrigatórias no Excel:
Código: Código da empresa.
Empresa: Nome da empresa.
Contato Onvio: Nome do contato no Onvio.
Grupo Onvio: Nome do grupo no Onvio.
Valor da Parcela: Valor da parcela pendente.
Data de Vencimento: Data de vencimento da parcela.
Carta de Aviso: Número do lembrete (1 a 6).


Mensagens: Usa templates "Cobranca_1" a "Cobranca_6" de mensagens.json.

3. ComuniCertificado

Objetivo: Lembra clientes de renovar certificados digitais.
Colunas Obrigatórias no Excel:
Codigo: Código da empresa.
Empresa: Nome da empresa.
Contato Onvio: Nome do contato no Onvio.
Grupo Onvio: Nome do grupo no Onvio.
CNPJ: CNPJ da empresa.
Vencimento: Data de vencimento do certificado.
Carta de Aviso: Número do lembrete (1 a 4).


Mensagens: Usa templates "Certificado_1" a "Certificado_4" de mensagens.json.

4. ALL

Objetivo: Envio de mensagens genéricas.
Colunas Obrigatórias no Excel:
Codigo: Código da empresa.
EMPRESAS: Nome da empresa.
CONTATO ONVIO: Nome do contato no Onvio.
GRUPO ONVIO: Nome do grupo no Onvio.


Mensagem: Usa o template "Mensagem Padrão" de mensagens.json.

Exemplo de mensagens.json
{
  "Mensagem Padrão": "Teste Desconsiderando mensagem",
  "Prorrogação Contrato": "Prezado cliente,\nEspero que estejam bem.\n\nGostaríamos de informar que o contrato de experiência das seguintes pessoas está preste a vencer:\n\n{pessoas_vencimentos}\n\nPara darmos prosseguimento aos devidos registros, solicitamos a gentileza de nos confirmar se haverá prorrogação do contrato ou se ele será encerrado nesta data.\n\nAtenciosamente,\nEquipe DP - C&S.",
  "Cobranca_1": "Prezado cliente,\nNotamos que o pagamento referente aos nossos serviços contábeis da empresa: {nome}, conforme abaixo, ainda não foi registrado.\n\n{parcelas}\nTotal: R$ {total}\n\nAtenciosamente,\nEquipe Financeiro Canella & Santos.",
  "Certificado_1": "Prezado Cliente,\nEstamos entrando em contato para lembrá-lo que o certificado digital da sua empresa {nome} (CNPJ {cnpj_formatado}) está próximo do vencimento.\n\nAtenciosamente,\nAna Caroline - Controle e Gerenciamento"
}

Uso

Execute o Aplicativo:
python ONE.py


Interface Gráfica:

Escolher Modelo: Selecione o modelo desejado no menu suspenso.
Escolher Excel: Clique em "Selecionar Excel" e escolha o arquivo Excel.
Definir Linha Inicial: Insira a linha inicial (padrão: 2, para pular o cabeçalho).
Escolher Mensagem: Selecione um template de mensagem no menu suspenso.
Adicionar/Editar Mensagem: Clique em "Adicionar/Editar Mensagem" para criar ou modificar templates.
Remover Mensagem: Exclua templates desnecessários.
Iniciar Chrome de Automação: Clique em "Iniciar Chrome de Automação" para abrir o Chrome e fazer login no Onvio Messenger.
Iniciar Processamento: Clique em "Iniciar Processamento" para começar o envio de mensagens.
Cancelar Processamento: Clique em "Cancelar Processamento" para interromper.
Fechar Programa: Clique em "Fechar Programa" para sair.
Abrir Log: Clique em "Abrir Log" para visualizar o arquivo de log gerado.


Logs:

Os logs são salvos automaticamente em ~/Documents/AutoMessenger_Logs/automessenger_one_log_<timestamp>.txt.
O log registra todas as ações, sucessos e erros durante o processamento.



Observações

Formato do Excel: Certifique-se de que o arquivo Excel contém as colunas exatas exigidas pelo modelo selecionado.
Autenticação no Onvio: Faça login no Onvio Messenger antes de iniciar o processamento, usando o botão "Iniciar Chrome de Automação".
Conexão com a Internet: Necessária para automação com Selenium.
Erros Comuns:
Excel inválido: Verifique se as colunas correspondem ao modelo selecionado.
Falha na automação: Certifique-se de que o Chrome está atualizado e o perfil de usuário está configurado corretamente.


Segurança: O script usa um perfil específico do Chrome (C:\PerfisChrome\automacao\ stylized-textured-backgroundProfile 1). Evite usar este perfil para outras atividades.

Suporte
Para dúvidas ou problemas, entre em contato com a Equipe de TI da Canella & Santos.

