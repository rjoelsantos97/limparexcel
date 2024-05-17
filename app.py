import streamlit as st
import pandas as pd
import json
from streamlit_lottie import st_lottie
from io import BytesIO
from openpyxl import load_workbook

# Função para carregar animação Lottie de um arquivo
def load_lottiefile(filepath: str):
    with open(filepath, "r") as f:
        return json.load(f)

# Função para limpar os dados removendo linhas que contêm "Total"
def clean_data(df):
    df_clean = df[~df.apply(lambda row: row.astype(str).str.contains('Total').any(), axis=1)]
    return df_clean

# Função para salvar DataFrame em um arquivo Excel mantendo o formato
def save_to_excel(df, original_file, sheet_name):
    with BytesIO() as output:
        # Carregar o workbook original
        book = load_workbook(original_file)
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            writer.book = book
            writer.sheets = {ws.title: ws for ws in book.worksheets}
            
            # Escrever a folha limpa no mesmo lugar
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            writer.save()
            processed_data = output.getvalue()
        
    return processed_data

# Carregar animação Lottie do arquivo
lottie_loading = load_lottiefile("loading_animation.json")

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
            # Centralizar a animação e a mensagem
            st.markdown(
                """
                <style>
                .css-1cpxqw2 {
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                }
                </style>
                """,
                unsafe_allow_html=True,
            )
            
            # Mostrar animação durante o processamento
            st_lottie(lottie_loading, height=300, key="loading")
            
            # Mostrar mensagem personalizada
            st.markdown("<h2>A fazer magia... só demora um bocadinho, obrigado por esperar</h2>", unsafe_allow_html=True)
            
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
