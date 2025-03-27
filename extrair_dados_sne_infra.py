import os
import pdfplumber
import pandas as pd
import re

# Diretório onde estão os arquivos PDF
pasta_arquivos = r"C:\Users\Paola\Downloads\Testes QA\arquivos_pdfs"  # Altere para o caminho correto

# Lista para armazenar os dados extraídos
dados = []

# Percorre todos os arquivos na pasta
for arquivo in os.listdir(pasta_arquivos):
    if arquivo.lower().endswith(".pdf"):  # Verifica se é um arquivo PDF
        caminho_arquivo = os.path.join(pasta_arquivos, arquivo)

        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = "\n".join(pagina.extract_text() for pagina in pdf.pages if pagina.extract_text())
            # Expressão regular para encontrar a PLACA (7 caracteres alfanuméricos)
            placa_match = re.search(r"PLACA ESPÉCIE PAÍS\n?([A-Z0-9]{7})", texto_completo)

            # Expressão regular para encontrar o Número do AIT (padrão alfanumérico maior que 8 caracteres)
            ait_match = re.search(r"IDENTIFICAÇÃO DO AUTO DE INFRAÇÃO.*?(\b[A-Z0-9]{8,}\b)", texto_completo, re.DOTALL)

            placa = placa_match.group(1) if placa_match else "Não encontrado"
            ait = ait_match.group(1) if ait_match else "Não encontrado"

            # Adiciona os dados na lista
            dados.append([arquivo, placa, ait])

# Criando um DataFrame
df = pd.DataFrame(dados, columns=["Arquivo", "Placa", "Número do AIT"])

# Salvando em Excel
caminho_excel = os.path.join(pasta_arquivos, "dados_extraidos.xlsx")
df.to_excel(caminho_excel, index=False)

print(f"Extração concluída! Arquivo salvo em: {caminho_excel}")
