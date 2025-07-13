from langchain_openai import AzureChatOpenAI
from dotenv import load_dotenv
import os
import pandas as pd
import pdfplumber
from pypdf import PdfReader, PdfWriter
from tkinter import Tk, filedialog
from langchain.prompts import PromptTemplate
from io import StringIO
import warnings
warnings.filterwarnings("ignore")

Tk().withdraw() 
pdf_path = filedialog.askopenfilename(
    title="Selecione um arquivo PDF",
    filetypes=[("Arquivos PDF", "*.pdf")]
)

if not pdf_path:
    print("Nenhum arquivo foi selecionado. Encerrando o programa.")
    exit()


output_dir = "paginas"
os.makedirs(output_dir, exist_ok=True)

reader = PdfReader(pdf_path)
total_paginas = len(reader.pages)
paginas_por_arquivo = 1

for i in range(0, total_paginas, paginas_por_arquivo):
    writer = PdfWriter()
    for j in range(i, min(i + paginas_por_arquivo, total_paginas)):
        writer.add_page(reader.pages[j])

    num_arquivo = (i // paginas_por_arquivo) + 1
    output_filename = f"pagina_{num_arquivo}.pdf"
    output_path = os.path.join(output_dir, output_filename)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)
    print(f"Arquivo {output_filename} salvo com páginas {i+1} até {min(i + paginas_por_arquivo, total_paginas)}")


load_dotenv()
pasta_pdf = "paginas"

dados_todos = []

llm = AzureChatOpenAI(
    openai_api_key=os.getenv("OPENAI_API_KEY"),
    azure_endpoint=os.getenv("OPENAI_API_BASE"),
    deployment_name=os.getenv("DEPLOYMENT_NAME"),
    openai_api_version=os.getenv("OPENAI_API_VERSION"),
    temperature=0,
)


template = '''
Você é um assistente de extração de dados. Leia atentamente o texto a seguir e extraia **TODOS OS CONJUNTOS** de dados que encontrar. 
Podem existir múltiplos registros no mesmo texto. Para cada registro, extraia os seguintes campos:

- CPF
- Recebedor
- Data
- Quantia


Retorne **todos os registros encontrados** no formato **CSV delimitado por ;**, com o **cabeçalho uma única vez** 
seguido de uma linha para cada conjunto de dados. 

Não resuma, não agrupe e não ignore duplicatas. Mesmo que os dados sejam parecidos, extraia todos os conjuntos individualmente.

**IMPORTANTE:** não inclua ```csv ou qualquer outro marcador de bloco de código no início ou no fim.

Texto:
{text}
'''


prompt = PromptTemplate(input_variables=['text'], template=template)
chain = prompt | llm

for nome_arquivo in sorted(os.listdir(pasta_pdf)):
    if nome_arquivo.endswith(".pdf"):
        caminho_pdf = os.path.join(pasta_pdf, nome_arquivo)
        print(f"Lendo {nome_arquivo}...")

       
        text = ''
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ''

        if text.strip():
            try:
                print("Extraindo com IA...")
                response = chain.invoke({'text': text}).content 
                df_temp = pd.read_csv(StringIO(response), sep=';')
                df_temp['arquivo'] = nome_arquivo  
                dados_todos.append(df_temp)
            except Exception as e:
                print(f"Erro ao processar {nome_arquivo}: {e}")
        else:
            print(f"Nenhum texto extraído de {nome_arquivo}")

if dados_todos:
    df_final = pd.concat(dados_todos, ignore_index=True)
    df_final = df_final.drop_duplicates()
    df_final.to_excel("Relatorio de Guia.xlsx", index=False)
    print("Relatório salvo com sucesso!")
else:
    print("Nenhum dado foi extraído.")

