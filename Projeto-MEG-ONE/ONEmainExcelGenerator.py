import os
import re
import pandas as pd
import pdfplumber
import openpyxl
from datetime import datetime, date
from collections import defaultdict
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from pathlib import Path

# Configuração do tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Função para ler o Excel de contatos
def carregar_contatos_excel(caminho_excel):
    contatos_dict = {}
    wb = openpyxl.load_workbook(caminho_excel)
    sheet = wb.active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if len(row) >= 4:
            codigo, nome, nome_contato, nome_grupo = row[:4]
            contatos_dict[str(codigo)] = {
                'empresa': nome,
                'contato': nome_contato,
                'grupo': nome_grupo
            }
    return contatos_dict

# Funções de processamento para cada modelo
def processar_one(pasta_pdf, excel_entrada, excel_saida, log_callback, progress_callback):
    codigos_empresas = []
    padrao = r'^(\d+)-'
    pdf_files = [f for f in os.listdir(pasta_pdf) if f.lower().endswith('.pdf')]
    log_callback(f"Encontrados {len(pdf_files)} arquivos PDF")
    progress_callback(0.2)

    for arquivo in pdf_files:
        match = re.match(padrao, arquivo)
        if match:
            codigo = match.group(1)
            codigos_empresas.append((codigo, arquivo))
            log_callback(f"Código encontrado: {codigo} - {arquivo}")
    
    progress_callback(0.4)
    log_callback("Lendo Excel de Contatos Onvio...")
    df_excel = pd.read_excel(excel_entrada)
    if df_excel.shape[1] < 4:
        raise ValueError("O arquivo Excel deve ter pelo menos 4 colunas (A-D).")
    df_excel.iloc[:, 0] = df_excel.iloc[:, 0].astype(str)
    
    progress_callback(0.6)
    log_callback("Comparando códigos e criando resultados...")
    resultados = []
    for codigo, arquivo_pdf in codigos_empresas:
        resultado = {
            'Código': codigo,
            'Empresa': '',
            'Contato Onvio': '',
            'Grupo Onvio': '',
            'Caminho': arquivo_pdf
        }
        if codigo in df_excel.iloc[:, 0].values:
            linha = df_excel[df_excel.iloc[:, 0] == codigo].iloc[0]
            resultado.update({
                'Empresa': linha.iloc[1],
                'Contato Onvio': linha.iloc[2],
                'Grupo Onvio': linha.iloc[3]
            })
            log_callback(f"Correspondência encontrada para código {codigo}")
        else:
            log_callback(f"Código {codigo} não encontrado no Excel")
        resultados.append(resultado)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(resultados)

def verifica_certificado_cobranca(data_vencimento):
    hoje = date.today()
    dias_passados = (hoje - data_vencimento).days
    if dias_passados <= 6:
        return 1
    elif dias_passados <= 14:
        return 2
    elif dias_passados <= 19:
        return 3
    elif dias_passados <= 24:
        return 4
    elif dias_passados <= 30:
        return 5
    else:
        return 6

def processar_cobranca(caminho_pdf, excel_entrada, excel_saida, log_callback, progress_callback):
    contatos_dict = carregar_contatos_excel(excel_entrada)
    log_callback("Lendo arquivo PDF...")
    progress_callback(0.2)
    
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text()
    
    linhas_texto = texto_completo.split('\n')
    regex_cliente = re.compile(r'Cliente: (\d+)')
    regex_nome = re.compile(r'Nome: (.+)')
    regex_parcela = re.compile(r'(\d{2}/\d{2}/\d{4}) (\d{1,3}(?:\.\d{3})*,\d{2})')
    
    dados = defaultdict(list)
    codigo_atual = None
    empresa_atual = None
    
    progress_callback(0.4)
    log_callback("Extraindo informações do PDF...")
    for linha in linhas_texto:
        match_cliente = regex_cliente.search(linha)
        if match_cliente:
            codigo_atual = str(match_cliente.group(1))
        match_nome = regex_nome.search(linha)
        if match_nome and codigo_atual:
            empresa_atual = match_nome.group(1)
        match_parcela = regex_parcela.search(linha)
        if match_parcela and codigo_atual and empresa_atual:
            data_vencimento = str(match_parcela.group(1))
            valor_parcela = round(float(match_parcela.group(2).replace(".", "").replace(",",".")), 2)
            data_venci = datetime.strptime(data_vencimento, '%d/%m/%Y').date()
            carta = verifica_certificado_cobranca(data_venci)
            contato_individual = contatos_dict.get(codigo_atual, {}).get('contato', '')
            contato_grupo = contatos_dict.get(codigo_atual, {}).get('grupo', '')
            dados[codigo_atual].append({
                'Código': codigo_atual,
                'Empresa': empresa_atual,
                'Contato Onvio': contato_individual,
                'Grupo Onvio': contato_grupo,
                'Valor da Parcela': valor_parcela,
                'Data de Vencimento': data_vencimento,
                'Carta de Aviso': carta
            })
    
    linhas = []
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df = pd.DataFrame(linhas)
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(linhas)

def processar_renovacao(excel_base, excel_entrada, excel_saida, log_callback, progress_callback):
    contatos_dict = carregar_contatos_excel(excel_entrada)
    log_callback("Lendo Excel Base...")
    progress_callback(0.2)
    
    df_comparacao = pd.read_excel(excel_base)
    codigos = df_comparacao.iloc[:, 0]
    pessoas = df_comparacao.iloc[:, 1]
    
    dados = {}
    progress_callback(0.4)
    log_callback("Comparando códigos e criando resultados...")
    for codigo_atual, pessoa in zip(codigos, pessoas):
        codigo_atual = str(codigo_atual)
        contato_individual = contatos_dict.get(codigo_atual, {}).get('contato', '')
        contato_grupo = contatos_dict.get(codigo_atual, {}).get('grupo', '')
        if codigo_atual not in dados:
            dados[codigo_atual] = []
        dados[codigo_atual].append({
            'Código': codigo_atual,
            'Empresa': pessoa,
            'Contato Onvio': contato_individual,
            'Grupo Onvio': contato_grupo
        })
    
    linhas = []
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df = pd.DataFrame(linhas)
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(linhas)

def formatar_cnpj(cnpj):
    cnpj_str = re.sub(r'\D', '', str(cnpj))
    if cnpj_str.endswith('.0'):
        cnpj_str = cnpj_str[:-2]
    if len(cnpj_str) == 13:
        cnpj_str = '0' + cnpj_str
    elif len(cnpj_str) == 12:
        cnpj_str = '00' + cnpj_str
    return cnpj_str.zfill(14)

def verifica_certificado_comunicado(data_vencimento):
    hoje = datetime.today()
    dias_restantes = (data_vencimento - hoje).days
    if dias_restantes == 0:
        return 3
    elif 0 < dias_restantes <= 5:
        return 2
    elif dias_restantes > 5:
        return 1
    elif dias_restantes < 0:
        return 4
    else:
        return 0

def processar_comunicado(excel_base, excel_entrada, excel_saida, log_callback, progress_callback):
    contatos_dict = carregar_contatos_excel(excel_entrada)
    log_callback("Lendo Excel Base...")
    progress_callback(0.2)
    
    df_comparacao = pd.read_excel(excel_base)
    codigos = df_comparacao.iloc[:, 0]
    empresas = df_comparacao.iloc[:, 1]
    cnpjs = df_comparacao.iloc[:, 2]
    vencimentos = df_comparacao.iloc[:, 4]
    situacoes = df_comparacao.iloc[:, 7]
    
    dados = {}
    progress_callback(0.4)
    log_callback("Comparando códigos e criando resultados...")
    for codigo_atual, empresa, cnpj, vencimento, situacao in zip(codigos, empresas, cnpjs, vencimentos, situacoes):
        codigo_atual = str(codigo_atual)
        if not pd.isna(cnpj):
            carta = verifica_certificado_comunicado(vencimento)
            cnpj_str = formatar_cnpj(cnpj)
            contato_individual = contatos_dict.get(codigo_atual, {}).get('contato', '')
            contato_grupo = contatos_dict.get(codigo_atual, {}).get('grupo', '')
            vencimento_str = vencimento.strftime("%d/%m/%Y") if isinstance(vencimento, pd.Timestamp) else str(vencimento)
            if codigo_atual not in dados:
                dados[codigo_atual] = []
            dados[codigo_atual].append({
                'Código': codigo_atual,
                'Empresa': empresa,
                'Contato Onvio': contato_individual,
                'Grupo Onvio': contato_grupo,
                'CNPJ': cnpj_str,
                'Vencimento': vencimento_str,
                'Carta de Aviso': carta
            })
    
    linhas = []
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)
    
    progress_callback(0.8)
    log_callback("Salvando arquivo Excel de saída...")
    df = pd.DataFrame(linhas)
    df.to_excel(excel_saida, index=False)
    log_callback(f"Arquivo Excel gerado com sucesso: {excel_saida}")
    return len(linhas)

# Mapeamento de modelos para funções de processamento
processadores = {
    "ONE": processar_one,
    "Cobranca": processar_cobranca,
    "ProrContrato": processar_renovacao,
    "ComuniCertificado": processar_comunicado
}

class ExcelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatórios Excel")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        self.pasta_pdf = ""
        self.excel_base = ""
        self.excel_entrada = ""
        self.excel_saida = ""
        self.modelo = ""
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ctk.CTkFrame(self.root, corner_radius=15)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        title_label = ctk.CTkLabel(
            main_frame, 
            text="🔗 Gerador de Relatórios Excel",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=(15, 20))
        
        fields_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        fields_frame.pack(fill="x", padx=20, pady=10)
        
        # Campo para selecionar modelo
        modelo_label = ctk.CTkLabel(
            fields_frame,
            text="Selecione o Modelo:",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        modelo_label.pack(anchor="w", pady=(0, 3))
        self.modelo_combobox = ctk.CTkComboBox(
            fields_frame,
            values=list(processadores.keys()),
            command=self.update_inputs,
            font=ctk.CTkFont(size=10),
            height=30
        )
        self.modelo_combobox.pack(fill="x", pady=6)
        
        # Frames para os campos de entrada
        self.inputs_frame = ctk.CTkFrame(fields_frame, fg_color="transparent")
        self.inputs_frame.pack(fill="x", pady=6)
        
        self.pdf_entry = None
        self.excel_base_entry = None
        self.input_entry = None
        self.output_entry = None
        
        # Botão processar
        self.process_button = ctk.CTkButton(
            main_frame,
            text="Processar Relatórios",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            command=self.process_files
        )
        self.process_button.pack(pady=(20, 10))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(main_frame, width=300)
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)
        
        # Label de status
        self.status_label = ctk.CTkLabel(
            main_frame,
            text="Selecione um modelo para começar",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        self.status_label.pack(pady=5)
        
        # Área de log
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(10, 15))
        
        log_title = ctk.CTkLabel(
            log_frame,
            text="Log de Processamento",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        log_title.pack(pady=(10, 5))
        
        self.log_text = ctk.CTkTextbox(
            log_frame,
            height=100,
            font=ctk.CTkFont(size=9)
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Rodapé
        footer_label = ctk.CTkLabel(
            main_frame,
            text="© 2025 SAFE v1.0 - Desenvolvido por Hugo",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        footer_label.pack(pady=5)
    
    def create_file_field(self, parent, label_text, button_text, command):
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=6)
        label = ctk.CTkLabel(
            field_frame,
            text=label_text,
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        label.pack(anchor="w", pady=(0, 3))
        input_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
        input_frame.pack(fill="x")
        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Nenhum arquivo selecionado",
            height=30,
            font=ctk.CTkFont(size=10)
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        button = ctk.CTkButton(
            input_frame,
            text=button_text,
            width=110,
            height=30,
            command=command
        )
        button.pack(side="right")
        return entry
    
    def update_inputs(self, choice):
        self.modelo = choice
        for widget in self.inputs_frame.winfo_children():
            widget.destroy()
        
        if choice == "ONE":
            self.pdf_entry = self.create_file_field(
                self.inputs_frame, 
                "Pasta com arquivos PDF:", 
                "Selecionar Pasta", 
                self.select_pdf_folder
            )
        elif choice == "Cobrança":
            self.pdf_entry = self.create_file_field(
                self.inputs_frame, 
                "Arquivo PDF:", 
                "Selecionar PDF", 
                self.select_pdf_file
            )
        else:
            self.excel_base_entry = self.create_file_field(
                self.inputs_frame, 
                "Excel Base:", 
                "Selecionar Excel", 
                self.select_excel_base
            )
        
        self.input_entry = self.create_file_field(
            self.inputs_frame, 
            "Excel de Contatos Onvio:", 
            "Selecionar Arquivo", 
            self.select_input_excel
        )
        self.output_entry = self.create_file_field(
            self.inputs_frame, 
            "Arquivo Excel de saída:", 
            "Definir Local", 
            self.select_output_excel
        )
        self.status_label.configure(text="Pronto para processar")
    
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta com arquivos PDF")
        if folder:
            self.pasta_pdf = folder
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, folder)
            self.log_message(f"Pasta PDF selecionada: {folder}")
    
    def select_pdf_file(self):
        file = filedialog.askopenfilename(
            title="Selecionar arquivo PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.pasta_pdf = file
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, file)
            self.log_message(f"PDF selecionado: {file}")
    
    def select_excel_base(self):
        file = filedialog.askopenfilename(
            title="Selecionar Excel Base",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.excel_base = file
            self.excel_base_entry.delete(0, "end")
            self.excel_base_entry.insert(0, file)
            self.log_message(f"Excel Base selecionado: {file}")
    
    def select_input_excel(self):
        file = filedialog.askopenfilename(
            title="Selecionar Excel de Contatos Onvio",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.excel_entrada = file
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, file)
            self.log_message(f"Excel de Contatos Onvio selecionado: {file}")
    
    def select_output_excel(self):
        file = filedialog.asksaveasfilename(
            title="Definir arquivo Excel de saída",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.excel_saida = file
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, file)
            self.log_message(f"Excel de saída definido: {file}")
    
    def log_message(self, message):
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.root.update_idletasks()
    
    def validate_inputs(self):
        if not self.modelo:
            messagebox.showerror("Erro", "Por favor, selecione um modelo.")
            return False
        if self.modelo == "ONE" and not self.pasta_pdf:
            messagebox.showerror("Erro", "Por favor, selecione a pasta com arquivos PDF.")
            return False
        if self.modelo == "Cobranca" and not self.pasta_pdf:
            messagebox.showerror("Erro", "Por favor, selecione o arquivo PDF.")
            return False
        if self.modelo in ["Prorcontrato", "ComuniCertificado"] and not self.excel_base:
            messagebox.showerror("Erro", "Por favor, selecione o Excel Base.")
            return False
        if not self.excel_entrada:
            messagebox.showerror("Erro", "Por favor, selecione o Excel de Contatos Onvio.")
            return False
        if not self.excel_saida:
            messagebox.showerror("Erro", "Por favor, defina o local do arquivo Excel de saída.")
            return False
        if self.modelo == "ONE" and not os.path.isdir(self.pasta_pdf):
            messagebox.showerror("Erro", "A pasta de PDFs não é válida.")
            return False
        if self.modelo == "Cobranca" and not os.path.isfile(self.pasta_pdf):
            messagebox.showerror("Erro", "O arquivo PDF não é válido.")
            return False
        if self.modelo in ["ProrContrato", "ComuniCertificado"] and not os.path.exists(self.excel_base):
            messagebox.showerror("Erro", "O Excel Base não existe.")
            return False
        if not os.path.exists(self.excel_entrada):
            messagebox.showerror("Erro", "O Excel de Contatos Onvio não existe.")
            return False
        return True
    
    def process_files(self):
        if not self.validate_inputs():
            return
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        try:
            self.progress_bar.set(0)
            self.status_label.configure(text="Processando...")
            self.log_text.delete("1.0", "end")
            self.log_message("Iniciando processamento...")
            
            processador = processadores.get(self.modelo)
            if not processador:
                raise ValueError(f"Modelo {self.modelo} não encontrado.")
            
            input_file = self.pasta_pdf if self.modelo in ["ONE", "Cobrança"] else self.excel_base
            total_registros = processador(input_file, self.excel_entrada, self.excel_saida, 
                                         self.log_message, self.progress_bar.set)
            
            self.progress_bar.set(1.0)
            self.status_label.configure(text="Processamento concluído!")
            self.log_message(f"Total de registros processados: {total_registros}")
            messagebox.showinfo("Sucesso", f"Processamento concluído!\nArquivo salvo em: {self.excel_saida}")
        
        except Exception as e:
            self.progress_bar.set(0)
            self.status_label.configure(text="Erro no processamento")
            self.log_message(f"ERRO: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")

def main():
    root = ctk.CTk()
    app = ExcelGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()