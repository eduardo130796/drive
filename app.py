import streamlit as st
import requests

# SEU TOKEN DE ACESSO
ACCESS_TOKEN = "ya29.a0AeXRPp4mfC_T4EaPeH0GtFB1E0qbzxeT4sJuoTuJTPyv03ulnuRJyHj-W2g40nS5PkFAoQAn8NQdM1kgxXWTzQF4D77w75m-Men7VHVv-buj_kijxyisgtV5h-c1ICbpkY9iIBkgXKmjgwy02m-K-yGN2TpwUOBjCawzXjmeaCgYKAdoSARMSFQHGX2MiSUQnRx8qz13c0h1nv-s9sg0175"

# ID da pasta específica
PASTA_ID = "1MP1wiSpSBx_pNTt3PTo3ayokZFxm_niR"  # Coloque o ID da sua pasta aqui

# Função para listar arquivos na pasta
def listar_arquivos(pasta_id):
    url = "https://www.googleapis.com/drive/v3/files"
    params = {
        "q": f"'{pasta_id}' in parents and trashed=false",  # Busca apenas arquivos na pasta especificada
        "fields": "files(id, name)"
    }
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}

    response = requests.get(url, params=params, headers=headers)
    arquivos = response.json().get("files", [])
    
    if arquivos:
        return arquivos
    else:
        return []

# Criando o aplicativo Streamlit
st.title("Listar Arquivos do Google Drive")

# Exibindo os arquivos encontrados
arquivos = listar_arquivos(PASTA_ID)

if arquivos:
    st.subheader("Arquivos encontrados na pasta:")
    for arquivo in arquivos:
        st.write(f"📄 {arquivo['name']} (ID: {arquivo['id']})")
else:
    st.write("❌ Nenhum arquivo encontrado na pasta.")
