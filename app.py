import streamlit as st
import pandas as pd

st.set_page_config(page_title="Analyseur Wallets LUIA", layout="wide")
st.title("📊 Analyse automatique de tokens LUIA")

# Upload CSV
tx_file = st.file_uploader("🧾 Importer le fichier CSV de transactions", type="csv")
holders_file = st.file_uploader("
