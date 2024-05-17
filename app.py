import streamlit as st
import pandas as pd
import json
from streamlit_lottie import st_lottie
from io import BytesIO
from openpyxl import load_workbook
import requests

# Função para carregar animação Lottie de uma URL
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

# Função para limpar os dados removendo linhas que contêm "Total"
def clean_data(df):
    df_clean = df[~df.apply(lambda row: row.astype(str).str.contains('Total').any(), axis=1)]
    return df_clean

# Função para salvar DataFrame em um arquivo Excel mantendo o formato
def save_to_excel(df, original_file, sheet_name):
    with BytesIO() as output:
        # Carregar o workbook original
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            book = load_workbook(original_file)
            writer.book = book
            writer.sheets = {ws.title: ws for ws in book.worksheets}
            
            # Escrever a folha limpa no mesmo lugar
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Salvar o arquivo modificado
            writer.save()
            processed_data = output.getvalue()
        
    return processed_data

# Carregar animação Lottie de uma URL
lottie_loading = load_lottieurl("https://assets7.lottiefiles.com/packages/lf20_zv4vcgk3.json")

# Título da aplicação
st.title("Upload de Arquivo Excel e Limpeza de Dados")

# Instrução para fazer upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

if uploaded_file is not None:
    # Ler o arquivo Excel
    df = pd.read_excel(uploaded_file, sheet_name=None)
    
    # Exibir os nomes das folhas do Excel
    sheet_names = df.keys()
    sheet_choice = st.selectbox("Escolha a folha para análise", list(sheet_names))
    
    # Ler a folha escolhida
    df_sheet = df[sheet_choice]
    
    # Exibir os dados originais
    st.subheader("Dados Originais")
    st.write(df_sheet)
    
    # Mostrar o botão para executar o tratamento
    if st.button("Executar Tratamento"):
        with st.spinner('Processando...'):
            # Mostrar animação durante o processamento
            st_lottie(lottie_loading, height=200, key="loading")
            
            # Limpar os dados
            df_clean = clean_data(df_sheet)
            
            # Exibir os dados limpos
            st.subheader("Dados Limpos (Sem Totais)")
            st.write(df_clean)
        
            # Salvar os dados limpos em um arquivo Excel
            processed_data = save_to_excel(df_clean, uploaded_file, sheet_choice)
        
        # Opção para baixar os dados limpos em formato Excel
        st.download_button(
            label="Baixar Dados Limpos",
            data=processed_data,
            file_name='dados_limpos.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
