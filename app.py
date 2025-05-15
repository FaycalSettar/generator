import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re
import traceback

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# =============================================
# SECTION 1: UPLOAD DES FICHIERS
# =============================================
with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation, Rép1 à Rép30)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    quizz_file = st.file_uploader("Fichier de correction (Quizz, .xlsx)", type="xlsx")

# =============================================
# SECTION 2: DÉTECTION DES QUESTIONS (inchangée)
# =============================================
def detecter_questions(doc):
    """Détection précise des questions avec regex améliorée"""
    questions = []
    current_question = None
    pattern = re.compile(r'^(\d+\.\d+)\s*[-–—)\s.]*\s*(.+?)\?$')
    reponse_pattern = re.compile(r'^([A-D])[\s\-–—).]+\s*(.*?)({{checkbox}})?\s*$')
   
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
       
        # Détection des questions
        match_question = pattern.match(texte)
        if match_question:
            current_question = {
                "index": i,
                "texte": f"{match_question.group(1)} - {match_question.group(2)}?",
                "reponses": [],
                "correct_idx": None
            }
            questions.append(current_question)
       
        # Détection des réponses
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1)
                texte_rep = match_reponse.group(2).strip()
                is_correct = match_reponse.group(3) is not None
               
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "correct": is_correct
                })
               
                if is_correct:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1
   
    return [q for q in questions if q["correct_idx"] is not None and len(q["reponses"]) >= 2]

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS (inchangée)
# =============================================
if word_file:
    if 'questions' not in st.session_state:
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}

    st.markdown("### Configuration des questions")
   
    for q in st.session_state.questions:
        q_id = q['index']
        q_num = q['texte'].split()[0]
       
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(
                f"Q{q_num}",
                value=st.session_state.figees.get(q_id, False),
                key=f"figer_{q_id}",
                help=q['texte']
            )
       
        with col2:
            if figer:
                options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                default_idx = q['correct_idx']
               
                bonne = st.selectbox(
                    f"Bonne réponse pour {q_num}",
                    options=options,
                    index=default_idx,
                    key=f"bonne_{q_id}"
                )
               
                st.session_state.figees[q_id] = True
                st.session_state.reponses_correctes[q_id] = options.index(bonne)

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION (modifiée)
# =============================================

def injecte_scores(doc, scores, total):
    mapping = {
        '{{result_mod1}}': str(scores[0]),
        '{{result_mod2}}': str(scores[1]),
        '{{result_mod3}}': str(scores[2]),
        '{{result_mod4}}': str(scores[3]),
        '{{result_mod5}}': str(scores[4]),
        '{{result_mod_total}}': str(total)
    }
    # Remplacement dans paragraphes
    for para in doc.paragraphs:
        for key, value in mapping.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, value)
    # Remplacement aussi dans les tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in mapping.items():
                    if key in cell.text:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.text = run.text.replace(key, value)

def score_reponses(reponses_utilisateur, correction):
    scores_par_module = [0, 0, 0, 0, 0]
    total = 0
    for i in range(1, 31):
        module_idx = (i - 1) // 6
        bonne_reponse = correction.get(str(i))
        reponse = reponses_utilisateur.get(f'Rép{i}')
        if str(reponse).strip().upper() == str(bonne_reponse).strip().upper():
            scores_par_module[module_idx] += 1
            total += 1
    return scores_par_module, total

def generer_document(row, template_path, correction):
    """Génération du QCM personnalisé avec scores injectés"""
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        for para in doc.paragraphs:
            for key, value in replacements.items():
                para.text = para.text.replace(key, value)
        # Scoring
        reponses_utilisateur = {f'Rép{i}': str(row[f'Rép{i}']) for i in range(1, 31)}
        scores, total = score_reponses(reponses_utilisateur, correction)
        injecte_scores(doc, scores, total)

        # (Tu gardes ton traitement de checkbox ici si tu veux aussi générer des QCM vierges)
        # ...

        return doc
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

# =============================================
# SECTION 5: GÉNÉRATION PRINCIPALE (modifiée)
# =============================================
if excel_file and word_file and quizz_file and st.session_state.get('questions'):
    if st.button("Générer les QCM avec scoring", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                # Lecture Excel et correction
                df = pd.read_excel(excel_file)
                quizz_df = pd.read_excel(quizz_file)
                correction = dict(zip(quizz_df['Question'].astype(str), quizz_df['BonneRéponse'].astype(str)))

                # Vérification Excel
                required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation'] + [f'Rép{i}' for i in range(1, 31)]
                missing = [col for col in required_cols if col not in df.columns]
                if missing:
                    st.error(f"Colonnes manquantes : {', '.join(missing)}")
                    st.stop()

                # Sauvegarde template
                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    f.write(word_file.getbuffer())

                # Création archive
                zip_path = os.path.join(tmpdir, "QCM_Generes.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                   
                    for idx, row in df.iterrows():
                        try:
                            doc = generer_document(row, template_path, correction)
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            doc.save(os.path.join(tmpdir, filename))
                            zipf.write(os.path.join(tmpdir, filename), filename)
                            progress_bar.progress((idx + 1) / len(df))
                        except Exception as e:
                            st.error(f"Échec pour {row['Prénom']} {row['Nom']} : {str(e)}")
                            continue

                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.success("✅ Génération terminée avec succès !")
                    st.download_button(
                        "📥 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Personnalises.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"ERREUR CRITIQUE : {str(e)}")
                st.text(traceback.format_exc())
