import pdfplumber

def ler_texto_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for i, pagina in enumerate(pdf.pages):
                texto = pagina.extract_text()
                if texto:
                    texto_completo += f"\n--- Página {i+1} ---\n"
                    texto_completo += texto
            print("Texto extraído do PDF:\n")
            print(texto_completo)
            return texto_completo
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")
        return None

# Caminho do PDF (substitua pelo caminho real do seu arquivo)
caminho_pdf = r"C:\Users\VM001\Desktop\Hugo\Relatório de Avisos de Vencimentos 04.06.pdf"
ler_texto_pdf(caminho_pdf)