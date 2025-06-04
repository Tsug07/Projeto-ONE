import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import pdfplumber
import pandas as pd
import re
from collections import defaultdict

# Função para ler o Excel de contatos
def carregar_contatos_excel(caminho_excel):
    contatos_dict = {}
    try:
        wb = openpyxl.load_workbook(caminho_excel)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) >= 4:
                codigo, nome, contato, grupo = row[:4]
                contatos_dict[str(codigo)] = {
                    'empresa': nome,
                    'contato': contato,
                    'grupo': grupo
                }
        return contatos_dict
    except Exception as e:
        print(f"Erro ao carregar Excel de contatos: {e}")
        return {}

# Função para extrair informações do PDF
def extrair_informacoes_pdf(caminho_pdf, contatos_dict):
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo += texto + "\n"
    
    # Dividir o texto por linhas e limpar
    linhas_texto = [linha.strip() for linha in texto_completo.split('\n') if linha.strip()]
    
    dados = []
    codigo_atual = None
    empresa_atual = None
    
    # Expressões regulares corrigidas
    regex_empresa = re.compile(r'Empresa:\s*(\d+)\s*-\s*(.+)')
    regex_cnpj = re.compile(r'CNPJ:\s*([\d\.\/\-]+)')
    
    # Padrões para diferentes tipos de eventos
    patterns_eventos = [
        # Padrão para Vencimento de Férias com limite
        re.compile(r'(\d+)\s+(.+?)\s+Vencimento de 2º Férias\s+(\d{2}/\d{2}/\d{4})\s*-\s*Limite\s+(\d{2}/\d{2}/\d{4})'),
        # Padrão para Aviso Prévio de rescisão
        re.compile(r'(\d+)\s+(.+?)\s+Aviso Prévio de rescisão\s+(\d{2}/\d{2}/\d{4})'),
        # Padrão para Contrato experiência
        re.compile(r'(\d+)\s+(.+?)\s+Contrato experiência (1º vencimento|prorrogação)\s+(\d{2}/\d{2}/\d{4})'),
        # Padrão para Aniversário colaboradores
        re.compile(r'(\d+)\s+(.+?)\s+Aniversário colaboradores\s+(\d{2}/\d{2}/\d{4})'),
        # # Padrão para Retorno de afastamento
        # re.compile(r'(\d+)\s+(.+?)\s+Retorno de afastamento de Doença\s+(\d{2}/\d{2}/\d{4})'),
        # # Padrão para Envio rescisão eSocial
        # re.compile(r'(\d+)\s+(.+?)\s+Envio rescisão eSocial\s+(\d{2}/\d{2}/\d{4})')
    ]
    
    i = 0
    while i < len(linhas_texto):
        linha = linhas_texto[i]
        
        # Identificar empresa
        match_empresa = regex_empresa.search(linha)
        if match_empresa:
            codigo_atual = match_empresa.group(1).strip()
            empresa_atual = match_empresa.group(2).strip()
            print(f"Empresa encontrada: {codigo_atual} - {empresa_atual}")
            i += 1
            continue
        
        # Pular linha de CNPJ
        if regex_cnpj.search(linha):
            i += 1
            continue
        
        # Tentar extrair eventos
        if codigo_atual:
            evento_encontrado = False
            
            for pattern in patterns_eventos:
                match = pattern.search(linha)
                if match:
                    codigo_empregado = match.group(1)
                    colaborador = match.group(2).strip()
                    
                    # Determinar tipo de evento e data
                    if "Vencimento de 2º Férias" in linha:
                        evento = "Vencimento de 2º Férias"
                        data = match.group(3)
                        prazo = match.group(4) if len(match.groups()) >= 4 else data
                    # elif "Aviso Prévio de rescisão" in linha:
                    #     evento = "Aviso Prévio de rescisão"
                    #     data = match.group(3)
                    #     prazo = data
                    elif "Contrato experiência" in linha:
                        tipo_contrato = match.group(3)
                        evento = f"Contrato experiência {tipo_contrato}"
                        data = match.group(4)
                        prazo = data
                    elif "Aniversário colaboradores" in linha:
                        evento = "Aniversário colaboradores"
                        data = match.group(3)
                        prazo = data
                    # elif "Retorno de afastamento" in linha:
                    #     evento = "Retorno de afastamento de Doença"
                    #     data = match.group(3)
                    #     prazo = data
                    # elif "Envio rescisão eSocial" in linha:
                    #     evento = "Envio rescisão eSocial"
                    #     data = match.group(3)
                    #     prazo = data
                    else:
                        # Fallback para outros eventos
                        evento = "Evento não especificado"
                        data = match.groups()[-1]
                        prazo = data
                    
                    # Obter informações do dicionário de contatos
                    if codigo_atual in contatos_dict:
                        contato = contatos_dict[codigo_atual]['contato']
                        grupo = contatos_dict[codigo_atual]['grupo']
                        empresa = contatos_dict[codigo_atual]['empresa']
                    else:
                        contato = ''
                        grupo = ''
                        empresa = empresa_atual if empresa_atual else ''
                    
                    # Criar registro
                    registro = {
                        'Código': codigo_atual,
                        'Empresa': empresa,
                        'Contato Onvio': contato,
                        'Grupo Onvio': grupo,
                        # 'Código Empregado': codigo_empregado,
                        'Colaborador': colaborador,
                        'Evento': evento,
                        # 'Data': data,
                        'Prazo': prazo
                    }
                    dados.append(registro)
                    # print(f"Evento extraído: {colaborador} - {evento} - {data}")
                    evento_encontrado = True
                    break
        
        i += 1
    
    print(f"Total de registros extraídos: {len(dados)}")
    return dados

# Função para gerar Excel
def gerar_excel(dados, caminho_excel):
    if not dados:
        print("Nenhum dado para gerar Excel")
        return
    
    df = pd.DataFrame(dados)
    df.to_excel(caminho_excel, index=False)
    print(f"Arquivo Excel criado: {caminho_excel}")

# Funções para interface gráfica
def selecionar_pdf():
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
    )
    entrada_pdf.delete(0, tk.END)
    entrada_pdf.insert(0, caminho_pdf)

def selecionar_destino_excel():
    caminho_excel = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")),
        title="Salvar arquivo Excel"
    )
    entrada_excel.delete(0, tk.END)
    entrada_excel.insert(0, caminho_excel)

def selecionar_lista_contatos():
    caminho_contatos = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
    )
    entrada_contatos.delete(0, tk.END)
    entrada_contatos.insert(0, caminho_contatos)

# Função para processar o PDF e gerar o Excel
def processar():
    caminho_pdf = entrada_pdf.get()
    caminho_excel = entrada_excel.get()
    caminho_contatos = entrada_contatos.get()
    
    if not caminho_pdf or not caminho_excel:
        messagebox.showwarning("Erro", "Por favor, selecione o arquivo PDF e o local para salvar o Excel.")
        return
    
    try:
        # Carregar contatos se fornecido
        contatos_dict = {}
        if caminho_contatos:
            contatos_dict = carregar_contatos_excel(caminho_contatos)
        
        # Extrair dados do PDF
        linhas_extraidas = extrair_informacoes_pdf(caminho_pdf, contatos_dict)
        
        if linhas_extraidas:
            gerar_excel(linhas_extraidas, caminho_excel)
            messagebox.showinfo("Sucesso", f"O arquivo Excel foi gerado com sucesso!\n{len(linhas_extraidas)} registros extraídos.")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado foi extraído do PDF. Verifique o formato do arquivo.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")
        print(f"Erro detalhado: {e}")

# Interface gráfica
def main():
    global entrada_pdf, entrada_excel, entrada_contatos
    janela = tk.Tk()
    janela.title("Gerador do Excel ONE")
    janela.geometry("600x450")

    # Título
    titulo = tk.Label(janela, text="Extrator de Dados PDF para Excel", font=("Arial", 14, "bold"))
    titulo.pack(pady=10)

    # PDF
    frame_pdf = tk.Frame(janela)
    frame_pdf.pack(pady=10, padx=20, fill='x')
    
    lbl_pdf = tk.Label(frame_pdf, text="Arquivo PDF de Vencimentos:", font=("Arial", 10))
    lbl_pdf.pack(anchor='w')
    
    entrada_pdf = tk.Entry(frame_pdf, width=60)
    entrada_pdf.pack(side='left', padx=(0, 10), fill='x', expand=True)
    
    btn_pdf = tk.Button(frame_pdf, text="Selecionar PDF", command=selecionar_pdf)
    btn_pdf.pack(side='right')

    # Contatos (opcional)
    frame_contatos = tk.Frame(janela)
    frame_contatos.pack(pady=10, padx=20, fill='x')
    
    lbl_contatos = tk.Label(frame_contatos, text="Excel de Contatos (opcional):", font=("Arial", 10))
    lbl_contatos.pack(anchor='w')
    
    entrada_contatos = tk.Entry(frame_contatos, width=60)
    entrada_contatos.pack(side='left', padx=(0, 10), fill='x', expand=True)
    
    btn_contatos = tk.Button(frame_contatos, text="Selecionar Contatos", command=selecionar_lista_contatos)
    btn_contatos.pack(side='right')

    # Excel de saída
    frame_excel = tk.Frame(janela)
    frame_excel.pack(pady=10, padx=20, fill='x')
    
    lbl_excel = tk.Label(frame_excel, text="Salvar Excel como:", font=("Arial", 10))
    lbl_excel.pack(anchor='w')
    
    entrada_excel = tk.Entry(frame_excel, width=60)
    entrada_excel.pack(side='left', padx=(0, 10), fill='x', expand=True)
    
    btn_excel = tk.Button(frame_excel, text="Escolher Local", command=selecionar_destino_excel)
    btn_excel.pack(side='right')

    # Botão processar
    btn_processar = tk.Button(janela, text="Gerar Excel", command=processar, 
                             font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", 
                             pady=10, padx=20)
    btn_processar.pack(pady=20)

    # Informações
    info_text = tk.Label(janela, text="O arquivo de contatos é opcional. Se não fornecido,\napenas os dados básicos serão extraídos.", 
                        font=("Arial", 9), fg="gray")
    info_text.pack(pady=10)

    janela.mainloop()

if __name__ == '__main__':
    main()