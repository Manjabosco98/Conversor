import fitz  # PyMuPDF
import pandas as pd

def converter_pdf_para_excel(caminho_pdf, caminho_excel):
    """
    Converte um arquivo PDF para um arquivo Excel.

    Args:
        caminho_pdf (str): O caminho para o arquivo PDF de entrada.
        caminho_excel (str): O caminho para o arquivo Excel de saída.
    """
    try:
        # Abre o PDF
        documento = fitz.open(caminho_pdf)
        
        # Lista para armazenar DataFrames
        lista_de_dfs = []
        
        # Itera sobre cada página
        for pagina_num in range(len(documento)):
            pagina = documento.load_page(pagina_num)
            texto = pagina.get_text("text")
            
            # Processa o texto para DataFrame
            linhas = texto.split('\n')
            dados = [linha.split() for linha in linhas]
            df = pd.DataFrame(dados)
            
            lista_de_dfs.append(df)
        
        # Combina todos os DataFrames em um único DataFrame
        df_completo = pd.concat(lista_de_dfs, ignore_index=True)
        
        # Salva o DataFrame combinado como um arquivo Excel
        df_completo.to_excel(caminho_excel, index=False)
        
        print(f"Arquivo PDF convertido com sucesso para Excel: {caminho_excel}")
    except Exception as e:
        print(f"Erro ao converter arquivo: {e}")
