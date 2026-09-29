import streamlit as st
import io
from docx import Document
import pandas as pd

st.set_page_config(page_title="Le Joker Fiche V3", layout="wide", page_icon="🇸🇳")
st.markdown('<div style="background:linear-gradient(90deg,#00853F,#FDEF42,#E31B23);padding:14px;border-radius:10px;text-align:center;font-weight:bold">🇸🇳 LE JOKER FICHE - V3 FINALE - 3 GUIDES OFFICIELS</div>', unsafe_allow_html=True)

st.sidebar.header("Le Joker Fiche")
classe = st.sidebar.selectbox("Classe", ["CP - Le present 1er groupe", "CE1-CE2 Multigrade", "CE1", "CE2", "CM1", "CM2"])
discipline = st.sidebar.selectbox("Discipline", ["Conjugaison", "Grammaire", "Communication orale"])
lecon = st.sidebar.text_input("Lecon", "Le present des verbes du 1er groupe")
effectif = st.sidebar.number_input("Effectif", 10, 80, 32)

if st.sidebar.button("GENERER LE JOKER FICHE", type="primary"):
    st.session_state['go']=True

if 'go' not in st.session_state:
    st.info("Clique a gauche pour generer. En-tete = Le Joker Fiche")
    st.stop()

st.subheader(f"Le Joker Fiche : {lecon} - {classe}")
# ... tableau avec Q Maitre / R Eleves / CE2 tuteurs ...
