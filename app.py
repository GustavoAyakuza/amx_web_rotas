import streamlit as st
import os
import pandas as pd
from processadores.script_rotas import process_excel
from processadores.script_roteirizador import merge_routes

# Configuração inicial
st.set_page_config(page_title="Amx Roteirizador", layout="centered")
st.title("📦 Roteirizador Automático - AMX Route Planner")

# Diretórios de arquivos
TEMP_DIR = "arquivos/temporarios"
RESULT_DIR = "arquivos/resultados"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

# Upload do Excel bruto
st.header("1️⃣ Envie o Excel de Rotas do Cliente")
excel_bruto = st.file_uploader("Excel Bruto (.xlsx)", type=["xlsx"], key="bruto")

if excel_bruto:
    caminho_bruto = os.path.join(TEMP_DIR, "excel_bruto.xlsx")
    caminho_rts = os.path.join(TEMP_DIR, "RTS.xlsx")
    caminho_saida_controle = os.path.join(TEMP_DIR, "saida_controle.xlsx")

    with open(caminho_bruto, "wb") as f:
        f.write(excel_bruto.read())

    process_excel(caminho_bruto, caminho_saida_controle, caminho_rts)

    st.success("✅ Arquivos gerados com sucesso!")
    with open(caminho_rts, "rb") as f:
        st.download_button("⬇️ Baixar planilha RTS (para subir no Zeo)", f, file_name="RTS.xlsx")
    with open(caminho_saida_controle, "rb") as f:
        st.download_button("⬇️ Baixar planilha Saída Controle", f, file_name="saida_controle.xlsx")

    st.markdown("---")
    st.info("⚠️ Agora suba a planilha **RTS** no Zeo Route Planner. Após gerar as rotas e baixar o Excel do Zeo, envie os dois arquivos abaixo para continuar.")

    st.header("2️⃣ Envie os arquivos para gerar a Saída Controle Atualizada")
    zeo_file = st.file_uploader("Excel com as Rotas (gerado pelo Zeo)", type=["xlsx"], key="zeo")
    saida_controle_file = st.file_uploader("Planilha de Saída Controle (a original gerada acima)", type=["xlsx"], key="controle")

    if zeo_file and saida_controle_file:
        caminho_zeo = os.path.join(TEMP_DIR, "zeo.xlsx")
        caminho_saida_controle_retorno = os.path.join(TEMP_DIR, "saida_controle_entrada.xlsx")
        caminho_saida_final = os.path.join(RESULT_DIR, "saida_controle_atualizada.xlsx")

        with open(caminho_zeo, "wb") as f:
            f.write(zeo_file.read())
        with open(caminho_saida_controle_retorno, "wb") as f:
            f.write(saida_controle_file.read())

        merge_routes(caminho_saida_controle_retorno, caminho_zeo, caminho_saida_final)
        st.success("✅ Saída Controle Atualizada gerada com sucesso!")
        with open(caminho_saida_final, "rb") as f:
            st.download_button("⬇️ Baixar Saída Controle Atualizada", f, file_name="saida_controle_atualizada.xlsx")
