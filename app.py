import streamlit as st
import os, io
from docx import Document

st.set_page_config(page_title="Le Joker - Agent Pédagogique Primaire", layout="wide", page_icon="🇸🇳")

# CLE SIMPLE - ON IGNORE LE FICHIER POUR LE MOMENT
CLE = "JOKER-DAKAR-2026"

if "unlock" not in st.session_state:
    st.session_state.unlock = False

if not st.session_state.unlock:
    st.title("🔐 Le Joker - Accès Protégé")
    k = st.text_input("Clé API", type="password", placeholder="Tape JOKER")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Déverrouiller"):
            if k.strip().upper() in ["JOKER", "JOKER-DAKAR-2026", CLE]:
                st.session_state.unlock = True
                st.rerun()
            else:
                st.error("Clé invalide mon ami - Tape JOKER")
    with c2:
        if st.button("🚨 BYPASS - Ouvrir sans clé"):
            st.session_state.unlock = True
            st.rerun()
    st.info("Tape JOKER en majuscule ou clique BYPASS")
    st.stop()

# BANDEAU BLEU QUE TU VOULAIS
st.markdown("""
<div style="background:#2a5bd7;color:white;padding:18px;border-radius:10px;text-align:center">
<h2 style="margin:0;color:white">Le Joker - Agent Pédagogique Primaire</h2>
<p style="margin:0">Déverrouillé 🔓 | CEB Sénégal Officiel</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1,1.4])
with col1:
    st.subheader("⚙️ Paramètres")
    classe=st.selectbox("Classe",["CI","CP","CE1","CE2","CM1","CM2"], index=0)
    discipline=st.selectbox("Discipline",["Mathématiques - Nombres et calculs","Français - Grammaire","ESVS"])
    notion=st.text_input("Notion *","Faire du vélo")
    os=st.text_area("OS","A la fin, l'élève doit être capable de faire du vélo")
    ia=st.text_input("IA","Dakar")
    ief=st.text_input("IEF","Pikine Guédiawaye")
    ecole=st.text_input("École","Dalifort")
    effectif=st.number_input("Effectif",10,120,45)
    duree=st.selectbox("Durée",["30 min","45 min","60 min"])

with col2:
    st.subheader(f"Fiche CEB - {classe} - {notion}")
    if st.button("🇸🇳 Générer Fiche CEB Complète", type="primary"):
        st.markdown(f"**REPUBLIQUE DU SENEGAL** - IA:{ia} | IEF:{ief} | Ecole:{ecole} | Classe:{classe} | Eff:{effectif}")
        st.markdown(f"**Leçon:** {notion} | **OS:** {os} | **Durée:** {duree}")

        st.table([
            ["PHASE","Étapes","Activités Maître","Activités Élève","Durée"],
            ["I. INTRO","Révision","Rappelle prérequis","Répond","5 min"],
            ["","Motivation",f"Situation {notion} marché","Observe","3 min"],
            ["","Annonce OS",os,"Écoute","2 min"],
            ["II. DEV","Présentation",f"Montre {notion}","Observe","10 min"],
            ["","Analyse","Travail groupe","Cherche","15 min"],
            ["","Synthèse","Règle","Copie","10 min"],
            ["III. CONCL","Évaluation","Exos cahier","Fait","10 min"],
            ["","Remédiation","Corrige","Corrige","3 min"],
            ["","Prolongement","Devoirs","Note","2 min"],
        ])

        # FICHIER WORD VRAI
        doc = Document()
        doc.add_heading("REPUBLIQUE DU SENEGAL - MEN", 2)
        doc.add_paragraph(f"IA: {ia} | IEF: {ief} | Ecole: {ecole} | Classe: {classe} | Eff: {effectif} | Leçon: {notion} | Durée: {duree}\nOS: {os}\nDiscipline: {discipline}")
        doc.add_heading("DEROULEMENT CEB", 3)
        table = doc.add_table(rows=1, cols=5); table.style='Table Grid'
        h=table.rows[0].cells; h[0].text='PHASE'; h[1].text='Étapes'; h[2].text='Maître'; h[3].text='Élève'; h[4].text='Durée'
        for r in [["I.INTRO","Révision","Rappelle","Répond","5 min"],["","Motivation","Situation","Observe","3 min"],["","OS",os,"Écoute","2 min"],["II.DEV","Présentation",notion,"Observe","10 min"],["","Analyse","Groupe","Cherche","15 min"],["","Synthèse","Règle","Copie","10 min"],["III.CONCL","Eval","Exos","Fait","10 min"],["","Remed","Corrige","Corrige","3 min"],["","Devoirs","Donne","Note","2 min"]]:
            row=table.add_row().cells
            for i in range(5): row[i].text=r[i]
        bio=io.BytesIO(); doc.save(bio)
        st.download_button("📥 TÉLÉCHARGER WORD.DOCX (s'ouvre direct)", bio.getvalue(), file_name=f"Fiche_CEB_{classe}_{notion}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.balloons()