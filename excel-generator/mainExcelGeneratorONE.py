import os
import re
import pandas as pd

# Caminho da pasta com os arquivos PDF
pasta_pdf = r"C:\Users\VM001\Documents\Relatorios"  # Substitua pelo caminho da pasta com PDFs

# Caminho do arquivo Excel de entrada
excel_entrada = r"C:\Users\VM001\Documents\HUGO\getContatoMessenger.xlsx"  # Substitua pelo caminho do Excel

# Caminho do arquivo Excel de saída
excel_saida = r"C:\Users\VM001\Documents\HUGO\ONE_relatorios_01.07.xlsx"

# Lista para armazenar os códigos das empresas a partir dos PDFs
codigos_empresas = []

# Expressão regular para capturar o número antes do hífen
padrao = r'^(\d+)-'

# Passo 1: Ler os códigos dos arquivos PDF
for arquivo in os.listdir(pasta_pdf):
    if arquivo.lower().endswith('.pdf'):
        match = re.match(padrao, arquivo)
        if match:
            codigo = match.group(1)  # Captura o número antes do hífen
            codigos_empresas.append((codigo, arquivo))  # Armazena o código e o nome do arquivo

# Passo 2: Ler o arquivo Excel
df_excel = pd.read_excel(excel_entrada)

# Verifica se o Excel tem pelo menos 4 colunas (A-D)
if df_excel.shape[1] < 4:
    raise ValueError("O arquivo Excel deve ter pelo menos 4 colunas (A-D).")

# Converte a coluna A (índice 0) para string para garantir compatibilidade
df_excel.iloc[:, 0] = df_excel.iloc[:, 0].astype(str)

# Passo 3: Comparar códigos e criar lista de resultados
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
    
    # Adiciona o resultado à lista, independentemente de ter correspondência
    resultados.append(resultado)

# Passo 4: Criar novo DataFrame com os resultados
df_resultado = pd.DataFrame(resultados)

# Passo 5: Salvar o resultado em um novo arquivo Excel
df_resultado.to_excel(excel_saida, index=False)

print(f"Arquivo Excel gerado com sucesso: {excel_saida}")