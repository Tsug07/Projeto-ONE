import os
import re
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from pathlib import Path

# Configuração do tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class ExcelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatórios Excel")
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        
        # Variáveis para armazenar os caminhos
        self.pasta_pdf = ""
        self.excel_entrada = ""
        self.excel_saida = ""
        
        self.setup_ui()
        
    def setup_ui(self):
        # Container principal
        main_frame = ctk.CTkFrame(self.root, corner_radius=15)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Título
        title_label = ctk.CTkLabel(
            main_frame, 
            text="Gerador de Relatórios Excel",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=(15, 20))
        
        # Container para os campos
        fields_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        fields_frame.pack(fill="x", padx=20, pady=10)
        
        # Campo 1: Pasta PDF
        self.create_file_field(
            fields_frame, 
            "Pasta com arquivos PDF:", 
            "Selecionar Pasta", 
            self.select_pdf_folder,
            0
        )
        
        # Campo 2: Excel de entrada
        self.create_file_field(
            fields_frame, 
            "Excel de Contatos Onvio:", 
            "Selecionar Arquivo", 
            self.select_input_excel,
            1
        )
        
        # Campo 3: Excel de saída
        self.create_file_field(
            fields_frame, 
            "Arquivo Excel de saída:", 
            "Definir Local", 
            self.select_output_excel,
            2
        )
        
        # Botão processar
        process_button = ctk.CTkButton(
            main_frame,
            text="Processar Relatórios",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            command=self.process_files
        )
        process_button.pack(pady=(20, 10))
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(main_frame, width=300)
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)
        
        # Label de status
        self.status_label = ctk.CTkLabel(
            main_frame,
            text="Pronto para processar",
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
        
    def create_file_field(self, parent, label_text, button_text, command, row):
        # Frame para cada campo
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=6)
        
        # Label
        label = ctk.CTkLabel(
            field_frame,
            text=label_text,
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        label.pack(anchor="w", pady=(0, 3))
        
        # Frame para entrada e botão
        input_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
        input_frame.pack(fill="x")
        
        # Campo de entrada
        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Nenhum arquivo selecionado",
            height=30,
            font=ctk.CTkFont(size=10)
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        
        # Botão
        button = ctk.CTkButton(
            input_frame,
            text=button_text,
            width=110,
            height=30,
            command=command
        )
        button.pack(side="right")
        
        # Armazenar referência do entry
        if row == 0:
            self.pdf_entry = entry
        elif row == 1:
            self.input_entry = entry
        elif row == 2:
            self.output_entry = entry
    
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta com arquivos PDF")
        if folder:
            self.pasta_pdf = folder
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, folder)
            self.log_message(f"Pasta PDF selecionada: {folder}")
    
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
        if not self.pasta_pdf:
            messagebox.showerror("Erro", "Por favor, selecione a pasta com arquivos PDF.")
            return False
        
        if not self.excel_entrada:
            messagebox.showerror("Erro", "Por favor, selecione o Excel de Contatos Onvio.")
            return False
        
        if not self.excel_saida:
            messagebox.showerror("Erro", "Por favor, defina o local do arquivo Excel de saída.")
            return False
        
        if not os.path.exists(self.pasta_pdf):
            messagebox.showerror("Erro", "A pasta de PDFs não existe.")
            return False
        
        if not os.path.exists(self.excel_entrada):
            messagebox.showerror("Erro", "O Excel de Contatos Onvio não existe.")
            return False
        
        return True
    
    def process_files(self):
        if not self.validate_inputs():
            return
        
        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=self.run_processing)
        thread.daemon = True
        thread.start()
    
    def run_processing(self):
        try:
            self.progress_bar.set(0)
            self.status_label.configure(text="Processando...")
            
            # Limpar log
            self.log_text.delete("1.0", "end")
            self.log_message("Iniciando processamento...")
            
            # Lista para armazenar os códigos das empresas a partir dos PDFs
            codigos_empresas = []
            
            # Expressão regular para capturar o número antes do hífen
            padrao = r'^(\d+)-'
            
            # Passo 1: Ler os códigos dos arquivos PDF
            self.log_message("Lendo arquivos PDF...")
            self.progress_bar.set(0.2)
            
            pdf_files = [f for f in os.listdir(self.pasta_pdf) if f.lower().endswith('.pdf')]
            self.log_message(f"Encontrados {len(pdf_files)} arquivos PDF")
            
            for arquivo in pdf_files:
                match = re.match(padrao, arquivo)
                if match:
                    codigo = match.group(1)
                    codigos_empresas.append((codigo, arquivo))
                    self.log_message(f"Código encontrado: {codigo} - {arquivo}")
            
            self.progress_bar.set(0.4)
            
            # Passo 2: Ler o arquivo Excel
            self.log_message("Lendo Excel de Contatos Onvio...")
            df_excel = pd.read_excel(self.excel_entrada)
            
            # Verifica se o Excel tem pelo menos 4 colunas (A-D)
            if df_excel.shape[1] < 4:
                raise ValueError("O arquivo Excel deve ter pelo menos 4 colunas (A-D).")
            
            # Converte a coluna A (índice 0) para string para garantir compatibilidade
            df_excel.iloc[:, 0] = df_excel.iloc[:, 0].astype(str)
            
            self.progress_bar.set(0.6)
            
            # Passo 3: Comparar códigos e criar lista de resultados
            self.log_message("Comparando códigos e criando resultados...")
            resultados = []
            
            for codigo, arquivo_pdf in codigos_empresas:
                # Inicializa o dicionário com os dados do PDF
                resultado = {
                    'Código': codigo,
                    'Empresa': '',
                    'Contato Onvio': '',
                    'Grupo Onvio': '',
                    'Caminho': arquivo_pdf
                }
                
                # Verifica se o código do PDF está na coluna A (índice 0) do Excel
                if codigo in df_excel.iloc[:, 0].values:
                    # Obtém a linha correspondente do Excel
                    linha = df_excel[df_excel.iloc[:, 0] == codigo].iloc[0]
                    # Atualiza os campos com os dados do Excel
                    resultado.update({
                        'Empresa': linha.iloc[1],  # Coluna B (índice 1)
                        'Contato Onvio': linha.iloc[2],  # Coluna C (índice 2)
                        'Grupo Onvio': linha.iloc[3]  # Coluna D (índice 3)
                    })
                    self.log_message(f"Correspondência encontrada para código {codigo}")
                else:
                    self.log_message(f"Código {codigo} não encontrado no Excel")
                
                # Adiciona o resultado à lista
                resultados.append(resultado)
            
            self.progress_bar.set(0.8)
            
            # Passo 4: Criar novo DataFrame com os resultados
            df_resultado = pd.DataFrame(resultados)
            
            # Passo 5: Salvar o resultado em um novo arquivo Excel
            self.log_message("Salvando arquivo Excel de saída...")
            df_resultado.to_excel(self.excel_saida, index=False)
            
            self.progress_bar.set(1.0)
            self.status_label.configure(text="Processamento concluído!")
            self.log_message(f"Arquivo Excel gerado com sucesso: {self.excel_saida}")
            self.log_message(f"Total de registros processados: {len(resultados)}")
            
            # Mostrar mensagem de sucesso
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