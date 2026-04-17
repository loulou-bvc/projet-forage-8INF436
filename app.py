# app.py — Version visuelle de test (sans vrais modèles)
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Statut de menace", page_icon="🌿", layout="wide")
st.title("🌿 Prédiction du statut de menace d'une espèce")

# --- Saisie manuelle ---
st.header("📝 Saisir une espèce")

# Champs obligatoires
common  = st.text_input("Nom commun *", placeholder="Ex: Lion, Aigle royal, Rose sauvage...")
kingdom = st.radio("Règne *", ["🐘 Animal", "🌿 Plante", "🍄 Champignon"], horizontal=True)

class_name = st.selectbox("Type *", [
    # Animaux
    "Mammifère", "Oiseau", "Reptile", "Poisson",
    "Amphibien", "Insecte", "Crustacé", "Mollusque",
    # Plantes
    "Plante à fleurs", "Arbre", "Fougère", "Mousse",
    # Autres
    "Algue", "Champignon", "Autre"
])

# Champs optionnels dans un expander (cachés par défaut)
with st.expander("➕ Informations supplémentaires (optionnel)"):
    col1, col2 = st.columns(2)
    with col1:
        phylum = st.selectbox("Groupe principal", [
            "Vertébré (CHORDATA)",
            "Arthropode (ARTHROPODA)",
            "Mollusque (MOLLUSCA)",
            "Plante vasculaire (TRACHEOPHYTA)",
            "Autre"
        ])
    with col2:
        genus = st.text_input("Nom scientifique du genre", placeholder="Ex: Panthera")

st.caption("* Champs obligatoires")

if st.button("🔍 Prédire le statut de cette espèce"):

    # Vérification des champs obligatoires
    if not common:
        st.error("⚠️ Veuillez entrer le nom de l'espèce.")
    else:
        # Résultats simulés (à remplacer par les vrais modèles plus tard)
        resultats_simules = {
            "Random Forest": ("Menacé",     [0.05, 0.80, 0.15]),
            "KNN":           ("Menacé",     [0.10, 0.75, 0.15]),
            "SVM":           ("Non menacé", [0.60, 0.25, 0.15]),
        }

        st.subheader(f"🎯 Résultats pour : **{common}**")
        cols = st.columns(3)
        classes = ["Disparu", "Menacé", "Non menacé"]

        for i, (nom, (prediction, probas)) in enumerate(resultats_simules.items()):
            with cols[i]:
                couleur = "🔴" if prediction == "Menacé" else "🟢" if prediction == "Non menacé" else "⚫"
                st.metric(label=nom, value=f"{couleur} {prediction}")
                st.write("Probabilités :")
                for classe, proba in zip(classes, probas):
                    st.progress(proba, text=f"{classe}: {proba*100:.0f}%")

        # Métriques simulées
        st.subheader("📊 Performance des modèles")
        st.caption("Métriques calculées sur le jeu de test")
        df_metriques = pd.DataFrame({
            "Random Forest": {"Accuracy": "0.8821", "F1 Macro": "0.7634", "F1 Weighted": "0.8790"},
            "KNN":           {"Accuracy": "0.8102", "F1 Macro": "0.7201", "F1 Weighted": "0.8050"},
            "SVM":           {"Accuracy": "0.8540", "F1 Macro": "0.7412", "F1 Weighted": "0.8510"},
        }).T
        st.dataframe(df_metriques, use_container_width=True)