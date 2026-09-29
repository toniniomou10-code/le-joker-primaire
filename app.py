import streamlit as st
import pandas as pd

st.set_page_config(page_title="Le Joker Fiche V3", layout="wide", page_icon="🇸🇳")

st.markdown('<div style="background:linear-gradient(90deg,#00853F,#FDEF42,#E31B23);padding:15px;border-radius:12px;text-align:center;font-weight:900;color:black">🇸🇳 LE JOKER FICHE - V3 FINALE - 3 GUIDES OFFICIELS</div>', unsafe_allow_html=True)

st.sidebar.markdown("### Le Joker Fiche")
classe = st.sidebar.selectbox("Classe", ["CP", "CE1", "CE2", "CE1-CE2 Multigrade", "CM1", "CM2"])
discipline = st.sidebar.selectbox("Discipline", ["Conjugaison", "Grammaire", "Vocabulaire"])
lecon_input = st.sidebar.text_input("Lecon", "Le present des verbes du 1er groupe")
effectif = st.sidebar.number_input("Effectif", 10, 100, 32)

genere = st.sidebar.button("🔴 GÉNÉRER LE JOKER FICHE", type="primary")

if not genere and 'fiche_ok' not in st.session_state:
    st.info("👈 Clique à gauche sur GENERER")
    st.stop()

st.session_state['fiche_ok'] = True

st.title(f"Le Joker Fiche : {lecon_input} - {classe}")
st.markdown(f"Classe: {classe} | Eff: {effectif} | 45 min | {discipline} | Guide CEB p111")

data = {
    "Etapes": ["Prerequis (5 min)", "Situation Probleme (10 min)", "Construction regle (15 min)", "Fixation (10 min)", "Evaluation (5 min)"],
    "Activites": [
        "PLM ardoises: etre/avoir",
        "Contexte marche Guinguineo exigu - Moustapha Tine",
        "Modelage: Je chante tu chantes... terminaisons -e -es -e -ons -ez -ent",
        "Production differenciee",
        "Grille A/B/C regle 2/3"
    ],
    "Situations Differenciation Q/R": [
        "Q: Conjuguez etre? R CE1: Je suis / R CE2: Nous sommes au marche",
        "Contexte: Aminata vend mangues. Elle dit 'Je chanter'. Consigne CE1: souligne verbe. CE2: conjugue + justifie -ent. Reformulation par 1 CE1 et 1 CE2",
        "Binomes mixtes CE2 tuteur CE1 3 min. Trame: Au marche je chante, tu marchandes",
        "CE1 attendu: Je chante au marche. CE2 attendu: Au marche de Guinguineo, nous chantons car c'est Tabaski",
        "CE1: 2/3 reussies = A. CE2: 3 phrases avec donc alors car = A"
    ],
    "Techniques": ["PLM", "Collectif", "Modelage Binomes", "Differenciation", "Observation"],
    "Supports": ["Ardoises", "Image marche", "Corpus 5 verbes", "Banque mots", "Grille"]
}

df = pd.DataFrame(data)
st.dataframe(df, use_container_width=True)

st.markdown("### POINTS DE VIGILANCE INSPECTEUR")
st.markdown("1. Ne jamais meme tache 2 niveaux 2. CE2 tuteurs 3. Attentes differenciees 4. SSI obligatoire")

st.success("Fiche generee - En-tete: Le Joker Fiche")
