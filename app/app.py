import streamlit as st
import pandas as pd
import numpy as np
import joblib
import os
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="Credit Score Predictor",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -----------------------------------------------
# CSS - thème financier sombre
# -----------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"],
.main, .block-container {
    background-color: #0f1117 !important;
    color: #e8eaf0;
    font-family: 'DM Sans', sans-serif;
}
[data-testid="stHeader"] { background-color: #0f1117; }

.main-title {
    font-size: 2.2rem;
    font-weight: 600;
    color: #ffffff;
    letter-spacing: -0.02em;
    margin-bottom: 0.15rem;
}
.main-subtitle {
    font-size: 0.88rem;
    color: #6b7280;
    margin-bottom: 1.8rem;
}

.section-header {
    font-size: 0.7rem;
    font-weight: 600;
    color: #60a5fa;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin: 1.4rem 0 0.8rem 0;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid #1e2535;
}

.score-card {
    background: #161b27;
    border: 1px solid #1e2535;
    border-radius: 12px;
    padding: 1.2rem 1rem;
    text-align: center;
}
.score-model {
    font-size: 0.7rem;
    color: #6b7280;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}
.score-good     { font-size: 1.5rem; font-weight: 600; color: #34d399; }
.score-standard { font-size: 1.5rem; font-weight: 600; color: #fbbf24; }
.score-poor     { font-size: 1.5rem; font-weight: 600; color: #f87171; }

.vote-box {
    background: linear-gradient(135deg, #1a2540, #1e2a45);
    border: 1px solid #2d3e6b;
    border-radius: 14px;
    padding: 1.3rem 2rem;
    text-align: center;
    margin-top: 1rem;
}
.vote-label { font-size: 0.72rem; color: #6b7280; text-transform: uppercase; letter-spacing: 0.08em; }
.vote-value { font-size: 1.9rem; font-weight: 600; margin-top: 0.25rem; }

[data-testid="stButton"] button {
    font-family: 'DM Sans', sans-serif;
    font-weight: 500;
    border-radius: 8px;
}
[data-testid="stButton"] button[kind="primary"] {
    background: #2563eb;
    border: none;
    width: 100%;
    padding: 0.55rem 2rem;
}
[data-testid="stButton"] button[kind="primary"]:hover {
    background: #1d4ed8;
}

label, .stSelectbox label, .stNumberInput label, .stSlider label {
    color: #9ca3af !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
}

hr { border-color: #1e2535; }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------
# chargement modèles
# -----------------------------------------------
@st.cache_resource
def charger_modeles():
    base = os.path.dirname(os.path.abspath(__file__))
    for d in [os.path.join(base, 'models'), os.path.join(base, '..', 'models'), 'models']:
        if os.path.exists(d):
            models_dir = d
            break
    else:
        raise FileNotFoundError("Dossier models/ introuvable. Lance le notebook en entier d'abord.")
    for d in [os.path.join(base, 'data'), os.path.join(base, '..', 'data'), 'data']:
        if os.path.exists(d):
            data_dir = d
            break
    else:
        raise FileNotFoundError("Dossier data/ introuvable.")
    return {
        'scaler':       joblib.load(os.path.join(models_dir, 'scaler.pkl')),
        'imputer':      joblib.load(os.path.join(models_dir, 'imputer.pkl')),
        'encoders':     joblib.load(os.path.join(models_dir, 'encoders_cat.pkl')),
        'le_target':    joblib.load(os.path.join(models_dir, 'le_target.pkl')),
        'sc_pca':       joblib.load(os.path.join(models_dir, 'sc_pca.pkl')),
        'pca':          joblib.load(os.path.join(models_dir, 'pca.pkl')),
        'cols_sel':     joblib.load(os.path.join(models_dir, 'cols_sel.pkl')),
        'feature_cols': joblib.load(os.path.join(models_dir, 'feature_cols.pkl')),
        'rf':           joblib.load(os.path.join(models_dir, 'random_forest.pkl')),
        'svm':          joblib.load(os.path.join(models_dir, 'svm.pkl')),
        'gb':           joblib.load(os.path.join(models_dir, 'gradient_boosting.pkl')),
        'resultats':    pd.read_csv(os.path.join(data_dir, 'resultats_modeles.csv'), index_col=0),
    }


def preprocess_input(df_input, obj):
    df = df_input.copy()
    df.drop(columns=[c for c in ['ID', 'Customer_ID', 'Name', 'SSN', 'Month', 'Credit_Score'] if c in df.columns], inplace=True)

    for col in ['Age','Annual_Income','Monthly_Inhand_Salary','Num_Bank_Accounts',
                'Num_Credit_Card','Interest_Rate','Num_of_Loan','Delay_from_due_date',
                'Num_of_Delayed_Payment','Changed_Credit_Limit','Num_Credit_Inquiries',
                'Outstanding_Debt','Credit_Utilization_Ratio','Credit_History_Age',
                'Total_EMI_per_month','Amount_invested_monthly','Monthly_Balance']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('[^0-9.-]','',regex=True), errors='coerce')

    for col in obj['feature_cols']:
        if col not in df.columns:
            df[col] = np.nan
    df = df[obj['feature_cols']]

    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].fillna('inconnu')
        le = obj['encoders'].get(col)
        if le:
            df[col] = df[col].astype(str).apply(lambda v: le.transform([v])[0] if v in le.classes_ else 0)
        else:
            df[col] = 0

    df_imp    = pd.DataFrame(obj['imputer'].transform(df), columns=df.columns)
    df_scaled = pd.DataFrame(obj['scaler'].transform(df_imp), columns=df_imp.columns)
    return df_scaled


def predire(df_scaled, obj):
    le    = obj['le_target']
    X_sel = df_scaled[obj['cols_sel']]
    X_pca = obj['pca'].transform(obj['sc_pca'].transform(df_scaled))
    return {
        'Forêt Aléatoire':   le.inverse_transform(obj['rf'].predict(X_sel)),
        'SVM (LinearSVC)':   le.inverse_transform(obj['svm'].predict(X_pca)),
        'Gradient Boosting': le.inverse_transform(obj['gb'].predict(X_sel)),
    }


def score_card(score, model):
    cls = {'Good':'score-good','Standard':'score-standard','Poor':'score-poor'}.get(score,'score-standard')
    dot = {'Good':'●','Standard':'●','Poor':'●'}.get(score,'●')
    return f'<div class="score-card"><div class="score-model">{model}</div><div class="{cls}">{dot} {score}</div></div>'


def vote_card(score):
    color = {'Good':'#34d399','Standard':'#fbbf24','Poor':'#f87171'}.get(score,'#fff')
    return f'<div class="vote-box"><div class="vote-label">Verdict des 3 modèles</div><div class="vote-value" style="color:{color}">● {score}</div></div>'


# -----------------------------------------------
# INTERFACE
# -----------------------------------------------
st.markdown('<div class="main-title">💳 Credit Score Predictor</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="main-subtitle">8INF436 &nbsp;·&nbsp; Forage de données &nbsp;·&nbsp; '
    'Frantxa Cabrejos &nbsp;·&nbsp; Loup-Djabril Le Bivic &nbsp;·&nbsp; Nathan Razafindratsima</div>',
    unsafe_allow_html=True
)

try:
    obj = charger_modeles()
    ok  = True
except Exception as e:
    st.error(f"⚠️ {e}")
    st.info("Lance le notebook `projet_8INF436.ipynb` en entier pour générer les fichiers `.pkl` dans `models/`.")
    ok = False

if ok:
    tab1, tab2, tab3 = st.tabs(["✏️  Saisie manuelle", "📂  Import CSV", "📊  Métriques"])

    # ---- Onglet 1 : saisie manuelle ----
    with tab1:
        st.markdown('<div class="section-header">Informations du client</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)

        with c1:
            age           = st.number_input("Âge", 18, 100, 35)
            occupation    = st.selectbox("Profession", [
                'Scientist','Teacher','Engineer','Entrepreneur','Developer',
                'Lawyer','Media_Manager','Doctor','Journalist',
                'Manager','Accountant','Musician','Writer','Other'])
            annual_income = st.number_input("Revenu annuel (USD)", 0.0, value=50000.0, step=1000.0)

        with c2:
            monthly_salary    = st.number_input("Salaire mensuel net (USD)", 0.0, value=3500.0, step=100.0)
            num_bank_accounts = st.number_input("Nb comptes bancaires", 0, 20, 3)
            num_credit_cards  = st.number_input("Nb cartes de crédit", 0, 15, 2)

        with c3:
            outstanding_debt = st.number_input("Dette en cours (USD)", 0.0, value=800.0, step=100.0)
            delay_days       = st.number_input("Retard moyen paiement (jours)", 0, 60, 5)
            credit_util      = st.slider("Taux utilisation crédit (%)", 0.0, 100.0, 30.0)

        st.markdown("")
        if st.button("🔮  Prédire le score de crédit", type="primary"):
            data = {col: np.nan for col in obj['feature_cols']}
            data.update({
                'Age': age, 'Occupation': occupation,
                'Annual_Income': annual_income, 'Monthly_Inhand_Salary': monthly_salary,
                'Num_Bank_Accounts': num_bank_accounts, 'Num_Credit_Card': num_credit_cards,
                'Outstanding_Debt': outstanding_debt, 'Delay_from_due_date': delay_days,
                'Credit_Utilization_Ratio': credit_util,
            })
            df_scaled = preprocess_input(pd.DataFrame([data]), obj)
            preds     = predire(df_scaled, obj)

            st.markdown('<div class="section-header">Résultats</div>', unsafe_allow_html=True)
            cols = st.columns(3)
            for (nom, arr), col in zip(preds.items(), cols):
                with col:
                    st.markdown(score_card(arr[0], nom), unsafe_allow_html=True)

            votes = [v[0] for v in preds.values()]
            st.markdown(vote_card(max(set(votes), key=votes.count)), unsafe_allow_html=True)

    # ---- Onglet 2 : import CSV ----
    with tab2:
        st.markdown('<div class="section-header">Téléverser un fichier</div>', unsafe_allow_html=True)
        st.caption("Le fichier doit contenir les mêmes colonnes que le dataset d'entraînement.")

        fichier = st.file_uploader("Fichier CSV", type=['csv'], label_visibility="collapsed")
        if fichier:
            df_up = pd.read_csv(fichier, low_memory=False)
            st.caption(f"{df_up.shape[0]} lignes · {df_up.shape[1]} colonnes")
            st.dataframe(df_up.head(5), use_container_width=True)

            if st.button("🔮  Prédire pour tous les clients", type="primary"):
                with st.spinner("Prédiction en cours..."):
                    preds = predire(preprocess_input(df_up, obj), obj)

                df_res = pd.DataFrame({
                    'Forêt Aléatoire':   preds['Forêt Aléatoire'],
                    'SVM (LinearSVC)':   preds['SVM (LinearSVC)'],
                    'Gradient Boosting': preds['Gradient Boosting'],
                })
                st.success(f"✓ Prédictions générées pour {len(df_up)} clients.")
                st.dataframe(df_res, use_container_width=True)

    # ---- Onglet 3 : métriques ----
    with tab3:
        st.markdown('<div class="section-header">Résultats sur le jeu de test (20 %)</div>', unsafe_allow_html=True)
        df_m = obj['resultats']

        for nom in df_m.index:
            r = df_m.loc[nom]
            st.markdown(f"**{nom}**")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Accuracy",     f"{r['accuracy']:.3f}")
            m2.metric("Précision",    f"{r['precision']:.3f}")
            m3.metric("Rappel",       f"{r['recall']:.3f}")
            m4.metric("F1 weighted",  f"{r['f1']:.3f}")
            m5.metric("CV F1 (k=5)",  f"{r['cv_f1_mean']:.3f} ± {r['cv_f1_std']:.3f}")
            st.markdown("---")

        # graphique comparatif
        st.markdown('<div class="section-header">Comparaison visuelle</div>', unsafe_allow_html=True)

        fig, ax = plt.subplots(figsize=(9, 3.5))
        fig.patch.set_facecolor('#161b27')
        ax.set_facecolor('#161b27')

        x      = np.arange(len(df_m))
        width  = 0.2
        colors = ['#60a5fa', '#34d399', '#fbbf24', '#f87171']
        for i, (m, c) in enumerate(zip(['accuracy','precision','recall','f1'], colors)):
            ax.bar(x + i*width, df_m[m], width, label=m, color=c, alpha=0.85)

        ax.set_xticks(x + width*1.5)
        ax.set_xticklabels(df_m.index, color='#9ca3af', fontsize=9)
        ax.set_ylim(0, 1.05)
        ax.tick_params(colors='#6b7280')
        for spine in ax.spines.values():
            spine.set_color('#1e2535')
        ax.legend(fontsize=8, labelcolor='#9ca3af', facecolor='#1e2535', edgecolor='#2d3748')
        plt.tight_layout()
        st.pyplot(fig, use_container_width=True)

        best = df_m['f1'].idxmax()
        st.markdown(
            f"<div style='text-align:center;color:#6b7280;font-size:0.85rem;margin-top:0.6rem'>"
            f"Meilleur modèle · <span style='color:#60a5fa;font-weight:600'>{best}</span>"
            f" · F1 = {df_m.loc[best,'f1']:.4f}</div>",
            unsafe_allow_html=True
        )
