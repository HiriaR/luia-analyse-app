import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from datetime import datetime
import tempfile
import os

st.set_page_config(page_title="Analyseur Wallets LUIA", layout="wide")
st.title("📊 Analyse automatique de tokens LUIA")

# Upload
tx_file = st.file_uploader("🧾 Importer le fichier CSV de transactions", type="csv")
holders_file = st.file_uploader("📄 Importer le fichier CSV de holders", type="csv")

# Aliases
st.subheader("🔖 Associer un nom à un wallet")
raw_aliases = st.text_area("Noms personnalisés", placeholder="0xabc123: Kale\\n0xdef456: Hiria")

alias_map = {}
if raw_aliases:
    for line in raw_aliases.splitlines():
        if ":" in line:
            addr, name = line.split(":", 1)
            alias_map[addr.strip().lower()] = name.strip()

if tx_file and holders_file:
    st.success("✅ Fichiers chargés, analyse en cours...")

    transactions = pd.read_csv(tx_file)
    holders = pd.read_csv(holders_file)

    transactions['DateTime (UTC)'] = pd.to_datetime(transactions['DateTime (UTC)'])
    transactions['Quantity'] = transactions['Quantity'].astype(str).str.replace(',', '').astype(float)
    holders['Balance'] = holders['Balance'].astype(str).str.replace(',', '').astype(float)

    latest_time = transactions['DateTime (UTC)'].max()
    start_time = latest_time - pd.Timedelta(hours=24)
    recent_sales = transactions[transactions['DateTime (UTC)'] >= start_time]
    sold_by_wallet = recent_sales.groupby('From')['Quantity'].sum().reset_index()
    sold_by_wallet.columns = ['Wallet', 'Total_Vendu_24h']

    top_20 = holders.head(20)
    merged = top_20.merge(sold_by_wallet, left_on='HolderAddress', right_on='Wallet', how='left')

    # 🧽 Normaliser les adresses pour comparaison
    merged['HolderAddress'] = merged['HolderAddress'].str.lower()
    exclude_wallets = ["0x2200C5ac2f9D8d635dB040A6F4eAb5ef6BF9E855"]
    exclude_wallets = [addr.lower() for addr in exclude_wallets]

    # ✅ Filtrer PancakeSwap Pool et autres
    merged = merged[~merged['HolderAddress'].isin(exclude_wallets)]

    merged['Nom'] = merged['HolderAddress'].str.lower().map(alias_map).fillna("")
    merged['Label'] = merged['Nom']
    merged.loc[merged['Label'] == '', 'Label'] = merged['HolderAddress']

    st.subheader("📋 Résumé")
    st.dataframe(merged[['Nom', 'HolderAddress', 'Balance', 'Total_Vendu_24h']])

    # Graphe
    st.subheader("📈 Soldes vs Ventes")
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(merged['Label'], merged['Balance'], label='Solde actuel')
    ax.bar(merged['Label'], merged['Total_Vendu_24h'].fillna(0), label='Vendu (24h)', alpha=0.7)
    ax.set_xticks(range(len(merged)))
    ax.set_xticklabels(merged['Label'], rotation=90)
    ax.legend()
    st.pyplot(fig)

    # Export Word
    st.subheader("📤 Export Word du rapport")

    if st.button("📝 Générer et télécharger le rapport Word"):
        doc = Document()
        doc.add_heading("Analyse des ventes sur 24h – Token LUIA", level=1)
        doc.add_paragraph(f"Généré le : {datetime.utcnow().strftime('%d/%m/%Y %H:%M UTC')}")
        doc.add_paragraph("Ce rapport montre les mouvements des top holders sur les 24 dernières heures, en croisant les soldes actuels et les ventes réalisées.")

        # Enregistrer l'image temporairement
        chart_path = tempfile.mktemp(suffix=".png")
        fig.savefig(chart_path)
        doc.add_picture(chart_path, width=Inches(6.5))

        doc.add_heading("Détail des holders ayant vendu", level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Light List'
        hdr = table.rows[0].cells
        hdr[0].text = "Nom"
        hdr[1].text = "Wallet"
        hdr[2].text = "Solde"
        hdr[3].text = "Ventes 24h"

        for _, row in merged.iterrows():
            if pd.notna(row['Total_Vendu_24h']) and row['Total_Vendu_24h'] > 0:
                cells = table.add_row().cells
                cells[0].text = row['Nom']
                cells[1].text = row['HolderAddress']
                cells[2].text = f"{row['Balance']:,.2f}"
                cells[3].text = f"{row['Total_Vendu_24h']:,.2f}"

        doc.add_paragraph("\\nDocument généré automatiquement et signé par Jarvis.")
        output_path = tempfile.mktemp(suffix=".docx")
        doc.save(output_path)

        with open(output_path, "rb") as f:
            st.download_button("📄 Télécharger le rapport Word", f, file_name="rapport_luia.docx")

        os.remove(chart_path)
        os.remove(output_path)
