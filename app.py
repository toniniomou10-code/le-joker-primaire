import streamlit as st
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Le Joker V2 - CEB Officiel", layout="wide", page_icon="🇸🇳")

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
                st.error("Clé invalide - Tape JOKER")
    with c2:
        if st.button("🚨 BYPASS - Ouvrir sans clé"):
            st.session_state.unlock = True
            st.rerun()
    st.stop()

# DISCIPLINES CEB COMPLETES - CONFORME GUIDE REVISE
DISCIPLINES = {
    "Langue et Communication": [
        "Lecture / Décodage / Compréhension",
        "Écriture / Graphisme / Copie",
        "Grammaire",
        "Conjugaison",
        "Orthographe",
        "Vocabulaire",
        "Expression écrite / Production d'écrits",
        "Communication orale / Expression orale",
        "Poésie / Récitation"
    ],
    "Mathématiques": [
        "Numération / Nombres et calculs",
        "Mesure",
        "Géométrie",
        "Résolution de problèmes"
    ],
    "Éveil / Découverte du Monde": [
        "Découverte du Vivant (ESVS)",
        "Découverte de la Matière",
        "Espace - Temps",
        "Vivre Ensemble / ECM",
        "Science et Technologie"
    ],
    "Autres": [
        "EPS",
        "Éducation Artistique - Dessin",
        "Éducation Artistique - Chant / Musique"
    ]
}
ALL_DISC = []
for dom, lst in DISCIPLINES.items():
    for d in lst:
        ALL_DISC.append(f"{dom} - {d}")

st.markdown("""
<div style="background:#2a5bd7;color:white;padding:18px;border-radius:10px;text-align:center">
<h2 style="margin:0;color:white">Le Joker V2 - CEB Officiel</h2>
<p style="margin:0">Conforme Guide CEB Révisé Sénégal | Mode Inspecteur Exigeant</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1,1.6])
with col1:
    st.subheader("⚙️ Paramètres simplifiés")
    classe = st.selectbox("Classe *", ["CI","CP","CE1","CE2","CM1","CM2"], index=1)
    discipline_full = st.selectbox("Discipline CEB *", ALL_DISC, index=3)
    notion = st.text_input("Leçon / Notion *", "Le présent des verbes du 1er groupe")
    os_input = st.text_area("Objectif Spécifique (OS) *", "A la fin, l'élève doit être capable de conjuguer au présent")
    effectif = st.number_input("Effectif", 10, 120, 45)
    duree = st.selectbox("Durée", ["30 min","45 min","60 min"], index=1)
    materiel = st.text_input("Matériel", "Tableau, craies, ardoises, corpus")
    ecole = st.text_input("École (facultatif)", "")

def generer_fiche(discipline, notion, classe, os):
    is_conj = "Conjugaison" in discipline
    is_math = "Mathématiques" in discipline

    if is_conj:
        q_intro = "Qu'est-ce qu'un verbe? Donnez un exemple.\nHier nous avons vu le verbe chanter à l'infinitif. Comment on le reconnait?"
        r_intro = "C'est un mot qui dit ce qu'on fait. Ex: manger, courir.\nIl se termine par -er, verbe du 1er groupe."
        q_dev = f"Observez: Je chante, Tu chantes, Il chante... Que remarquez-vous à la fin?\nQui peut entourer la terminaison pour Je? Nous?\nSi je dis {notion}, quelle est la règle?"
        r_dev = "La fin change, le début reste.\nJe -> e, Tu -> es, Il -> e, Nous -> ons, Vous -> ez, Ils -> ent\nOn garde le radical + terminaison"
        app = "1. (Ardoises) Conjugue 'parler' au présent.\n2. (Cahier d'essai) Complète: Nous...... (danser) bien.\n3. (Binômes) Chacun conjugue un verbe et fait corriger."
        eval_c = "Consigne: Conjugue au présent.\na) Je (manger) une mangue.\nb) Vous (chanter) bien.\nc) Ils (jouer) au foot.\nCritères: -e/-es/-e/-ons/-ez/-ent corrects"
    elif is_math:
        q_intro = "Comptez de 2 en 2 jusqu'à 20.\nRappel addition"
        r_intro = "2,4,6,8... / Réponses élèves"
        q_dev = f"Situation: 3 sachets de 4 mangues. Combien en tout?\nComment calculer vite?"
        r_dev = "4+4+4 = 12\nC'est 3x4=12"
        app = "1. Calcule: 2x5, 3x4\n2. Problème: 4 tables de 6 élèves\n3. Dessine et calcule"
        eval_c = "Résous: a) 5x3=? b) Problème boutique. Critère: calcul juste"
    else:
        q_intro = f"Qu'avons-nous vu hier sur {notion}?\nQui peut donner un exemple?"
        r_intro = "Rappel acquis / Exemples élèves"
        q_dev = f"Observez ce corpus sur {notion}\nQue remarquez-vous?\nComment on explique la règle?"
        r_dev = "On voit que... / Formulation règle par élèves"
        app = f"1. Identification sur {notion}\n2. Transformation\n3. Production personnelle"
        eval_c = f"Exercice écrit sur {notion} avec 3 items gradués"

    return [
        ["Phase","Étapes & Durée","Objectif","Questions du Maître","Réponses Élèves + Activités","Supports"],
        ["I. INTRO (10 min)", "1. Révision\n(5 min)", "Vérifier acquis", q_intro, r_intro, "Tableau, ardoises"],
        ["", "2. Motivation\n(3 min)", "Susciter intérêt", f'Situation vécue sénégalaise sur "{notion}": au marché, à la maison...\n"Que voyez-vous?"', "Observent, décrivent, hypothèses", "Image, corpus"],
        ["", "3. Annonce OS\n(2 min)", "Clarifier attente", f'"{os}"\n"Répétez ce qu\'on va apprendre?"', "Répètent et reformulent OS", "Voix"],
        ["II. DEV (25 min)", "4. Présentation\n(7 min)", f"Découvrir {notion}", q_dev, "Observation active", "Corpus tableau"],
        ["", "5. Analyse\n(10 min)", f"Comprendre {notion}", "Travail groupe 4-5 élèves. Circule, relance:\n- Que constatez-vous?\n- Pourquoi?", "Travail groupes, manipulations, rapporteurs", "Ardoises, cahier essai"],
        ["", "6. Synthèse\n(8 min)", "Fixer règle", "Fait dégager règle avec élèves:\n- Quelle est la règle?\nInstitutionnalise au tableau.", f"Formulent règle. Copient leçon.\nEx: {r_dev[:100]}", "Tableau, cahier leçons"],
        ["III. EVAL (10 min)", "7. Application\n(5 min)", "Fixer acquisition", app, "Font sur ardoises puis cahiers. Auto-correction.", "Ardoises, cahiers"],
        ["", "8. Évaluation\n(5 min)", "Mesurer OS", eval_c, "Travail individuel écrit", "Cahier éval"],
        ["", "9. Remédiation\n(3 min)", "Corriger", "Correction collective. Dépassement pour forts, ré-explication pour faibles demain. Devoirs.", f"Corrigent. Notent devoirs: Apprendre règle + 3 phrases avec {notion}", "Cahiers"]
    ]

with col2:
    st.subheader(f"Fiche - {classe} - {notion}")
    if st.button("🇸🇳 GÉNÉRER FICHE CEB COMPLÈTE V2", type="primary"):
        if not notion or not os_input:
            st.error("OS et Notion obligatoires - Inspecteur exigeant!")
        else:
            fiche_data = generer_fiche(discipline_full, notion, classe, os_input)
            st.markdown(f"**École:** {ecole} | **Classe:** {classe} | **Eff:** {effectif} | **Durée:** {duree}")
            st.markdown(f"**Discipline:** {discipline_full} | **Leçon:** {notion}")
            st.markdown(f"**OS:** {os_input} | **Matériel:** {materiel}")
            st.divider()
            for row in fiche_data[1:]:
                with st.expander(f"{row[0]} - {row[1]}"):
                    st.write(f"🎯 {row[2]}")
                    st.write(f"👨🏫 MAÎTRE (Questions précises):\n{row[3]}")
                    st.write(f"👨‍🎓 ÉLÈVES (Réponses attendues):\n{row[4]}")
                    st.write(f"📦 {row[5]}")

            doc = Document()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run("REPUBLIQUE DU SENEGAL - MEN\nFICHE PEDAGOGIQUE CEB")
            run.bold = True
            doc.add_paragraph(f"Ecole: {ecole} | Classe: {classe} | Eff: {effectif} | Durée: {duree}\nDiscipline: {discipline_full} | Leçon: {notion}\nOS: {os_input}\nMatériel: {materiel}")
            table = doc.add_table(rows=1, cols=6)
            table.style='Table Grid'
            hdr = table.rows[0].cells
            for i,h in enumerate(["PHASE","Étapes","Objectif","Maître (Questions)","Élèves (Réponses)","Supports"]):
                hdr[i].text = h
            for r in fiche_data[1:]:
                row = table.add_row().cells
                for i in range(6):
                    row[i].text = r[i]
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button("📥 TÉLÉCHARGER WORD V2", bio.getvalue(), file_name=f"Fiche_V2_{classe}_{notion.replace(' ','_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.balloons()
