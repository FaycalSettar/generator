import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re
import traceback
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# =============================================
# SECTION 1: UPLOAD DES FICHIERS
# =============================================
with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")

# =============================================
# SECTION 2: DÉTECTION DES QUESTIONS (CORRIGÉE)
# =============================================
def detecter_questions(doc):
    """Détection précise des questions avec regex améliorée"""
    questions = []
    current_question = None
    # Regex plus flexible pour les numéros de question
    pattern = re.compile(r'^(\d+[\.\d]*)\s*[-–—)\s.]*\s*(.+?)\?$')
    # Support des lettres maj/min et différents séparateurs
    reponse_pattern = re.compile(r'^([A-Za-z])[\s\-–—).]+\s*(.*?)({{checkbox}})?\s*$')
   
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
       
        # Détection des questions
        match_question = pattern.match(texte)
        if match_question:
            current_question = {
                "index": i,
                "texte": f"{match_question.group(1)} - {match_question.group(2)}?",
                "reponses": [],
                "correct_idx": None,
                "original_text": texte  # Sauvegarder le texte original
            }
            questions.append(current_question)
       
        # Détection des réponses
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1).upper()  # Standardiser en majuscules
                texte_rep = match_reponse.group(2).strip()
                is_correct = match_reponse.group(3) is not None
               
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "correct": is_correct,
                    "original_text": texte
                })
               
                if is_correct:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1

    # Filtrer les questions valides avec au moins 2 réponses et réponse correcte
    return [q for q in questions if q["correct_idx"] is not None and len(q["reponses"]) >= 2]

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file:
    if 'questions' not in st.session_state or st.session_state.get('current_template') != word_file.name:
        # Réinitialiser la configuration si nouveau template
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}
        st.session_state.current_template = word_file.name

    st.markdown("### Configuration des questions")
   
    for q in st.session_state.questions:
        q_id = q['index']
        q_num = q['texte'].split()[0]
       
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(
                f"Q{q_num}",
                value=st.session_state.figees.get(q_id, False),
                key=f"figer_{q_id}_{word_file.name[:5]}",  # Clé unique par template
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
                    key=f"bonne_{q_id}_{word_file.name[:5]}"
                )
               
                st.session_state.figees[q_id] = True
                st.session_state.reponses_correctes[q_id] = options.index(bonne)

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION (CORRIGÉE)
# =============================================
def remplacer_placeholders(paragraph, replacements):
    """Remplace les placeholders en préservant la mise en forme"""
    for key, value in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def reinserer_checkbox(para, texte, checkbox):
    """Réinsère la checkbox avec mise en forme cohérente"""
    # Nettoyer le paragraphe existant
    para.clear()
    
    # Créer un premier run avec le texte
    run = para.add_run(texte)
    run.font.size = Pt(11)
    
    # Ajouter la checkbox à la fin
    run = para.add_run(f" {checkbox}")
    run.font.size = Pt(11)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT

def generer_document(row, template_path):
    """Génération avec gestion correcte des checkboxes"""
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }

        # Remplacement des variables
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    remplacer_placeholders(para, {key: value})

        # Traitement des questions
        for q in st.session_state.questions:
            reponses = q['reponses'].copy()
            is_figee = st.session_state.figees.get(q['index'], False)
           
            if is_figee:
                # Réponses figées
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
                reponse_correcte = reponses.pop(bonne_idx)
                reponses.insert(0, reponse_correcte)
            else:
                # Mélanger en conservant la bonne réponse
                correct_idx = next((i for i, r in enumerate(reponses) if r['correct']), None)
                if correct_idx is not None:
                    reponse_correcte = reponses.pop(correct_idx)
                    reponses.insert(0, reponse_correcte)
                random.shuffle(reponses)

            # Mise à jour du document avec les réponses
            for rep in reponses:
                idx = rep['index']
                checkbox = "☑" if reponses.index(rep) == 0 else "☐"
                
                # Reconstruction du texte original avec lettre et réponse
                texte_base = rep['original_text'].split(' ', 1)[0]  # Lettre de réponse
                texte_reponse = rep['texte']
                ligne_complete = f"{texte_base} - {texte_reponse}"
                
                reinserer_checkbox(doc.paragraphs[idx], ligne_complete, checkbox)

        return doc
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

# =============================================
# SECTION 5: GÉNÉRATION PRINCIPALE
# =============================================
if excel_file and word_file and st.session_state.get('questions'):
    if st.button("Générer les QCM", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                # Vérification Excel
                df = pd.read_excel(excel_file)
                required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
               
                if not all(col in df.columns for col in required_cols):
                    missing = [col for col in required_cols if col not in df.columns]
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
                    total = len(df)
                   
                    for idx, row in df.iterrows():
                        try:
                            doc = generer_document(row, template_path)
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            doc.save(os.path.join(tmpdir, filename))
                            zipf.write(os.path.join(tmpdir, filename), filename)
                            progress_bar.progress((idx + 1)/total, text=f"Génération en cours : {idx+1}/{total}")
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
