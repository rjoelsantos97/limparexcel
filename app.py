import streamlit as st
import pandas as pd

# Função para limpar os dados removendo linhas que contêm "Total"
def clean_data(df):
    df_clean = df[~df.apply(lambda row: row.astype(str).str.contains('Total').any(), axis=1)]
    return df_clean

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
    
    # Limpar os dados
    df_clean = clean_data(df_sheet)
    
    # Exibir os dados limpos
    st.subheader("Dados Limpos (Sem Totais)")
    st.write(df_clean)

    # Opção para baixar os dados limpos
    st.download_button(
        label="Baixar Dados Limpos",
        data=df_clean.to_csv(index=False).encode('utf-8'),
        file_name='dados_limpos.csv',
        mime='text/csv'
    )
