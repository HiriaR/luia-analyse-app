import streamlit as st
import pandas as pd

st.set_page_config(page_title="Analyseur Wallets LUIA", layout="wide")
st.title("📊 Analyse automatique de tokens LUIA")

# Upload CSV
tx_file = st.file_uploader("🧾 Importer le fichier CSV de transactions", type="csv")
holders_file = st.file_uploader("📄 Importer le fichier CSV de holders", type="csv")

# Champ texte pour associer des noms à des wallets
st.subheader("🔖 Associer un nom à un wallet (ex: 0xabc123: Mon pote)")
raw_aliases = st.text_area("Noms personnalisés", placeholder="0xabc123: Kale le créateur\\n0xdef456: Hiria le blanc")

# Création du dictionnaire des alias
alias_map = {}
if raw_aliases:
    for line in raw_aliases.splitlines():
        if ":" in line:
            addr, name = line.split(":", 1)
            alias_map[addr.strip().lower()] = name.strip()

# Analyse si les fichiers sont présents
if tx_file and holders_file:
    st.success("✅ Fichiers CSV chargés ! Analyse en cours...")

    transactions = pd.read_csv(tx_file)
    holders = pd.read_csv(holders_file)

    transactions['DateTime (UTC)'] = pd.to_datetime(transactions['DateTime (UTC)'])
    transactions['Quantity'] = transactions['Quantity'].astype(str).str.replace(',', '').astype(float)
    holders['Balance'] = holders['Balance'].astype(str).str.replace(',', '').astype(float)

    # Calcul des ventes sur 24h
    latest_time = transactions['DateTime (UTC)'].max()
    window_start = latest_time - pd.Timedelta(hours=24)
    recent_sales = transactions[transactions['DateTime (UTC)'] >= window_start]
    sold_by_wallet = recent_sales.groupby('From')['Quantity'].sum().reset_index()
    sold_by_wallet.columns = ['Wallet', 'Total_Vendu_24h']

    # Top 20 holders
    top_20 = holders.head(20)
    merged = top_20.merge(sold_by_wallet, left_on='HolderAddress', right_on='Wallet', how='left')
    merged['Nom'] = merged['HolderAddress'].str.lower().map(alias_map).fillna("")

    st.subheader("📋 Résumé des top 20 holders")
    st.dataframe(merged[['Nom', 'HolderAddress', 'Balance', 'Total_Vendu_24h']])

    # Graphique
    st.subheader("📈 Graphique des soldes vs ventes")
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(merged['HolderAddress'], merged['Balance'], label='Solde actuel')
    ax.bar(merged['HolderAddress'], merged['Total_Vendu_24h'].fillna(0), label='Vendu (24h)', alpha=0.7)
    ax.set_xticklabels(merged['Nom'], rotation=90)
    ax.legend()
    st.pyplot(fig)

    st.success("✅ Analyse terminée. Rapport prêt à exporter bientôt.")
