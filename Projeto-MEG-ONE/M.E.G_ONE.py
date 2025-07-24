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
from PIL import Image, ImageTk

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
            # Converter código para inteiro e depois para string para remover .0
            codigo_limpo = str(int(float(codigo))) if codigo is not None else ""
            contatos_dict[codigo_limpo] = {
                'empresa': nome,
                'contato': nome_contato,
                'grupo': nome_grupo
            }
    return contatos_dict

# Função auxiliar para limpar e padronizar códigos
def limpar_codigo(codigo):
    """Converte código para string limpa, removendo .0 e espaços"""
    if codigo is None or pd.isna(codigo):
        return ""
    try:
        # Se for float com .0, remove o .0
        if isinstance(codigo, float) and codigo.is_integer():
            return str(int(codigo))
        # Se for string, remove espaços e .0 no final
        codigo_str = str(codigo).strip()
        if codigo_str.endswith('.0'):
            codigo_str = codigo_str[:-2]
        return codigo_str
    except:
        return str(codigo).strip()

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
            codigo_atual = limpar_codigo(match_cliente.group(1))  # CORREÇÃO AQUI
            log_callback(f"Debug - Código extraído do PDF: '{codigo_atual}'")
        match_nome = regex_nome.search(linha)
        if match_nome and codigo_atual:
            empresa_atual = match_nome.group(1)
        match_parcela = regex_parcela.search(linha)
        if match_parcela and codigo_atual and empresa_atual:
            data_vencimento = str(match_parcela.group(1))
            valor_parcela = round(float(match_parcela.group(2).replace(".", "").replace(",",".")), 2)
            data_venci = datetime.strptime(data_vencimento, '%d/%m/%Y').date()
            carta = verifica_certificado_cobranca(data_venci)
            
            # CORREÇÃO: Debug para verificar busca no dicionário
            contato_info = contatos_dict.get(codigo_atual, {})
            log_callback(f"Debug - Buscando código '{codigo_atual}' no dicionário: {contato_info}")
            
            contato_individual = contato_info.get('contato', '')
            contato_grupo = contato_info.get('grupo', '')
            
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
    
    # CORREÇÃO: Log do dicionário de contatos para debug
    log_callback(f"Debug - Contatos carregados: {len(contatos_dict)} registros")
    log_callback(f"Debug - Primeiros 3 códigos do dicionário: {list(contatos_dict.keys())[:3]}")
    
    df_comparacao = pd.read_excel(excel_base)
    codigos = df_comparacao.iloc[:, 0]
    pessoas = df_comparacao.iloc[:, 1]
    
    dados = {}
    progress_callback(0.4)
    log_callback("Comparando códigos e criando resultados...")
    for codigo_atual, pessoa in zip(codigos, pessoas):
        codigo_atual = limpar_codigo(codigo_atual)  # CORREÇÃO AQUI
        log_callback(f"Debug - Código do Excel Base: '{codigo_atual}' (tipo: {type(codigo_atual)})")
        
        # CORREÇÃO: Debug para verificar busca no dicionário
        contato_info = contatos_dict.get(codigo_atual, {})
        log_callback(f"Debug - Buscando código '{codigo_atual}' no dicionário: {contato_info}")
        
        contato_individual = contato_info.get('contato', '')
        contato_grupo = contato_info.get('grupo', '')
        
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
    
    # CORREÇÃO: Log do dicionário de contatos para debug
    log_callback(f"Debug - Contatos carregados: {len(contatos_dict)} registros")
    log_callback(f"Debug - Primeiros 3 códigos do dicionário: {list(contatos_dict.keys())[:3]}")
    
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
        codigo_atual = limpar_codigo(codigo_atual)  # CORREÇÃO AQUI
        log_callback(f"Debug - Código do Excel Base: '{codigo_atual}' (tipo: {type(codigo_atual)})")
        
        if not pd.isna(cnpj):
            carta = verifica_certificado_comunicado(vencimento)
            cnpj_str = formatar_cnpj(cnpj)
            
            # CORREÇÃO: Debug para verificar busca no dicionário
            contato_info = contatos_dict.get(codigo_atual, {})
            log_callback(f"Debug - Buscando código '{codigo_atual}' no dicionário: {contato_info}")
            
            contato_individual = contato_info.get('contato', '')
            contato_grupo = contato_info.get('grupo', '')
            vencimento_str = vencimento.strftime("%d/%m/%Y") if isinstance(vencimento, pd.Timestamp) else str(vencimento)
            
            if codigo_atual not in dados:
                dados[codigo_atual] = []
            dados[codigo_atual].append({
                'Codigo': codigo_atual,
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
        self.root.title("M.E.G_ONE - Main Excel Generator ONE V1.0")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        
        self.pasta_pdf = ""
        self.excel_base = ""
        self.excel_entrada = ""
        self.excel_saida = ""
        self.modelo = ""
        
        self.setup_ui()
        
    def load_logo(self):
        """Carrega o logo se existir, procurando no diretório atual"""
        try:
            # Obtém o diretório onde o script está sendo executado
            script_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Lista de possíveis nomes e extensões para o logo
            logo_files = [
                "logo.png", "logo.jpg", "logo.jpeg", "logo.ico", "logo.gif",
                "Logo.png", "Logo.jpg", "Logo.jpeg", "Logo.ico", "Logo.gif",
                "LOGO.png", "LOGO.jpg", "LOGO.jpeg", "LOGO.ico", "LOGO.gif"
            ]
            
            for logo_file in logo_files:
                logo_path = os.path.join(script_dir, logo_file)
                if os.path.exists(logo_path):
                    print(f"Logo encontrado: {logo_path}")
                    image = Image.open(logo_path)
                    # Redimensiona o logo para um tamanho compacto
                    image = image.resize((32, 32), Image.Resampling.LANCZOS)
                    return ctk.CTkImage(light_image=image, dark_image=image, size=(32, 32))
            
            print("Nenhum logo encontrado nos formatos suportados")
            return None
            
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")
            return None
        
    def setup_ui(self):
        # Container principal compacto
        main_frame = ctk.CTkFrame(self.root, corner_radius=10)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header compacto
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent", height=50)
        header_frame.pack(fill="x", padx=15, pady=(10, 5))
        header_frame.pack_propagate(False)
        
        # Título com logo (se disponível)
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(expand=True, fill="x")
        
        logo_image = self.load_logo()
        if logo_image:
            logo_label = ctk.CTkLabel(title_frame, image=logo_image, text="")
            logo_label.pack(side="left", padx=(0, 8))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text="M.E.G_ONE",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.pack(side="left", anchor="w")
        
        # Seleção de modelo
        model_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        model_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkLabel(
            model_frame,
            text="Modelo:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=(0, 8))
        
        self.modelo_combobox = ctk.CTkComboBox(
            model_frame,
            values=list(processadores.keys()),
            command=self.update_inputs,
            width=200,
            height=28
        )
        self.modelo_combobox.pack(side="left")
        
        # Frame para inputs dinâmicos
        self.inputs_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        self.inputs_frame.pack(fill="x", padx=15, pady=5)
        
        # Controles inferiores
        controls_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        controls_frame.pack(fill="x", padx=15, pady=5)
        
        # Botão processar
        self.process_button = ctk.CTkButton(
            controls_frame,
            text="🚀 Processar Relatórios",
            font=ctk.CTkFont(size=12, weight="bold"),
            height=35,
            command=self.process_files
        )
        self.process_button.pack(fill="x", pady=(0, 5))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(controls_frame, height=8)
        self.progress_bar.pack(fill="x", pady=2)
        self.progress_bar.set(0)
        
        # Status
        self.status_label = ctk.CTkLabel(
            controls_frame,
            text="Selecione um modelo para começar",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        self.status_label.pack(pady=2)
        
        # Log compacto
        log_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        log_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        log_header = ctk.CTkFrame(log_frame, fg_color="transparent", height=30)
        log_header.pack(fill="x", padx=10, pady=(8, 0))
        log_header.pack_propagate(False)
        
        ctk.CTkLabel(
            log_header,
            text="📋 Log:",
            font=ctk.CTkFont(size=11, weight="bold")
        ).pack(side="left")
        
        ctk.CTkButton(
            log_header,
            text="Limpar",
            width=60,
            height=24,
            command=self.clear_log
        ).pack(side="right")
        
        # Área de log
        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(size=9),
            height=100
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        
        # Rodapé
        footer_label = ctk.CTkLabel(
            main_frame,
            text="© 2025 - Desenvolvido por Hugo",
            font=ctk.CTkFont(size=9),
            text_color="gray50"
        )
        footer_label.pack(pady=5)
        
        # Inicialização
        self.log_message("Sistema inicializado. Selecione um modelo para começar.")
    
    def create_compact_field(self, parent, label_text, button_text, command):
        """Cria um campo de entrada compacto"""
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=2)
        
        # Label
        label = ctk.CTkLabel(
            field_frame,
            text=label_text,
            font=ctk.CTkFont(size=10, weight="bold"),
            width=120,
            anchor="w"
        )
        label.pack(side="left", padx=(0, 5))
        
        # Entry
        entry = ctk.CTkEntry(
            field_frame,
            placeholder_text="Nenhum arquivo selecionado",
            height=26,
            font=ctk.CTkFont(size=9)
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Button
        button = ctk.CTkButton(
            field_frame,
            text=button_text,
            width=80,
            height=26,
            command=command
        )
        button.pack(side="right")
        
        return entry
    
    def update_inputs(self, choice):
        """Atualiza os campos de entrada baseado no modelo selecionado"""
        self.modelo = choice
        
        # Limpa campos anteriores
        for widget in self.inputs_frame.winfo_children():
            widget.destroy()
        
        # Cria campos específicos do modelo
        if choice == "ONE":
            self.pdf_entry = self.create_compact_field(
                self.inputs_frame, 
                "📁 Pasta PDF:", 
                "Selecionar", 
                self.select_pdf_folder
            )
        elif choice == "Cobranca":
            self.pdf_entry = self.create_compact_field(
                self.inputs_frame, 
                "📄 Arquivo PDF:", 
                "Selecionar", 
                self.select_pdf_file
            )
        else:
            self.excel_base_entry = self.create_compact_field(
                self.inputs_frame, 
                "📊 Excel Base:", 
                "Selecionar", 
                self.select_excel_base
            )
        
        # Campos comuns
        self.input_entry = self.create_compact_field(
            self.inputs_frame, 
            "📋 Contatos Onvio:", 
            "Selecionar", 
            self.select_input_excel
        )
        
        self.output_entry = self.create_compact_field(
            self.inputs_frame, 
            "💾 Saída Excel:", 
            "Definir", 
            self.select_output_excel
        )
        
        self.status_label.configure(text="✅ Pronto para processar")
        self.log_message(f"Modelo selecionado: {choice}")
    
    def clear_log(self):
        """Limpa o log"""
        self.log_text.delete("1.0", "end")
        self.log_message("Log limpo")
    
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta com arquivos PDF")
        if folder:
            self.pasta_pdf = folder
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, os.path.basename(folder))
            self.log_message(f"📁 Pasta selecionada: {folder}")
    
    def select_pdf_file(self):
        file = filedialog.askopenfilename(
            title="Selecionar arquivo PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.pasta_pdf = file
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, os.path.basename(file))
            self.log_message(f"📄 PDF selecionado: {os.path.basename(file)}")
    
    def select_excel_base(self):
        file = filedialog.askopenfilename(
            title="Selecionar Excel Base",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.excel_base = file
            self.excel_base_entry.delete(0, "end")
            self.excel_base_entry.insert(0, os.path.basename(file))
            self.log_message(f"📊 Excel Base: {os.path.basename(file)}")
    
    def select_input_excel(self):
        file = filedialog.askopenfilename(
            title="Selecionar Excel de Contatos Onvio",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file:
            self.excel_entrada = file
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, os.path.basename(file))
            self.log_message(f"📋 Contatos Onvio: {os.path.basename(file)}")
    
    def select_output_excel(self):
        file = filedialog.asksaveasfilename(
            title="Definir arquivo Excel de saída",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.excel_saida = file
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, os.path.basename(file))
            self.log_message(f"💾 Saída definida: {os.path.basename(file)}")
    
    def log_message(self, message):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        self.log_text.insert("end", formatted_message)
        self.log_text.see("end")
        self.root.update_idletasks()
    
    def validate_inputs(self):
        """Valida se todos os campos necessários foram preenchidos"""
        if not self.modelo:
            messagebox.showerror("Erro", "Selecione um modelo.")
            return False
        
        if self.modelo == "ONE" and not self.pasta_pdf:
            messagebox.showerror("Erro", "Selecione a pasta com arquivos PDF.")
            return False
        
        if self.modelo == "Cobranca" and not self.pasta_pdf:
            messagebox.showerror("Erro", "Selecione o arquivo PDF.")
            return False
        
        if self.modelo in ["ProrContrato", "ComuniCertificado"] and not self.excel_base:
            messagebox.showerror("Erro", "Selecione o Excel Base.")
            return False
        
        if not self.excel_entrada:
            messagebox.showerror("Erro", "Selecione o Excel de Contatos Onvio.")
            return False
        
        if not self.excel_saida:
            messagebox.showerror("Erro", "Defina o arquivo Excel de saída.")
            return False
        
        return True
    
    def process_files(self):
        """Inicia o processamento em thread separada"""
        if not self.validate_inputs():
            return
        
        self.process_button.configure(state="disabled")
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        """Executa o processamento"""
        try:
            self.progress_bar.set(0)
            self.status_label.configure(text="🔄 Processando...")
            self.log_message("🚀 Iniciando processamento...")
            
            processador = processadores.get(self.modelo)
            if not processador:
                raise ValueError(f"Modelo {self.modelo} não encontrado.")
            
            input_file = self.pasta_pdf if self.modelo in ["ONE", "Cobranca"] else self.excel_base
            total_registros = processador(
                input_file, 
                self.excel_entrada, 
                self.excel_saida, 
                self.log_message, 
                self.progress_bar.set
            )
            
            self.progress_bar.set(1.0)
            self.status_label.configure(text="✅ Processamento concluído!")
            self.log_message(f"🎉 Total de registros: {total_registros}")
            self.log_message("✅ Processamento finalizado!")
            
            messagebox.showinfo(
                "Sucesso", 
                f"Processamento concluído!\n\nTotal de registros: {total_registros}\n\nArquivo salvo em:\n{self.excel_saida}"
            )
        
        except Exception as e:
            self.progress_bar.set(0)
            self.status_label.configure(text="❌ Erro no processamento")
            self.log_message(f"❌ ERRO: {str(e)}")
            messagebox.showerror("Erro", f"Erro durante o processamento:\n{str(e)}")
        
        finally:
            self.process_button.configure(state="normal")

def main():
    root = ctk.CTk()
    app = ExcelGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()