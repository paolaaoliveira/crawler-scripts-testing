import os
import fitz  # PyMuPDF
import csv
import re

# Pasta onde estão os arquivos PDF
pasta = r"C:\\Users\\Paola\\Downloads\\Testes QA\\arquivos_pdfs"

# Função para extrair dados do PDF
def extrair_dados(pdf_path):
    with fitz.open(pdf_path) as doc:
        texto = " ".join(page.get_text() for page in doc).upper()  # Captura o texto completo e coloca em maiúsculo para uniformizar
        
        # Extração do nome do arquivo
        nome_arquivo = os.path.basename(pdf_path)
        
        # Extração do CPF e Prazo usando regex para maior precisão
        cpf = "N/A"
        prazo = "N/A"

        match_cpf = re.search(r'CPF DO CONDUTOR[\s:]*([\d\.\-]+)', texto)
        match_prazo = re.search(r'PRAZO PARA REALIZA.*?NOVO EXAME[\s\w]*([\d]{2}/[\d]{2}/[\d]{4})', texto)

        if match_cpf:
            cpf = match_cpf.group(1)
        if match_prazo:
            prazo = match_prazo.group(1)

        return nome_arquivo, cpf, prazo

# Criar arquivo CSV com os resultados
with open("resultados_exames.csv", mode="w", newline="", encoding="utf-8-sig") as csvfile:
    escritor = csv.writer(csvfile, delimiter=';')  # Delimitador adequado para abrir no Excel
    escritor.writerow(["Nome do Arquivo", "CPF", "Prazo para Novo Exame"])

    # Percorre todos os arquivos da pasta
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".pdf"):
            caminho_completo = os.path.join(pasta, arquivo)
            dados = extrair_dados(caminho_completo)
            escritor.writerow(dados)

print("Processo concluído! Resultados salvos em 'resultados_exames.csv'")