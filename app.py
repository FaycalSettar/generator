import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re
import traceback
from collections import defaultdict

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# =============================================
# SECTION 1: UPLOAD DES FICHIERS
# =============================================
with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    correction_file = st.file_uploader("Fichier de correction (colonnes B: Numéro question, C: Réponse)", type="xlsx")

# =============================================
# SECTION 2: DÉTECTION DES QUESTIONS ET CORRECTIONS
# =============================================
def charger_corrections(fichier):
    df = pd.read_excel(fichier, header=None, usecols="B,C", skiprows=1)
    df.columns = ['Question', 'Reponse']
    return {row['Question']: row['Reponse'] for _, row in df.iterrows()}

def detecter_questions(doc, corrections):
    questions = []
    current_question = None
    pattern = re.compile(r'^(\d+\.\d+)\s*[-–—)\s.]*\s*(.+?)\?$')
    reponse_pattern = re.compile(r'^([A-D])[\s\-–—).]+\s*(.*?)({{checkbox}})?\s*$')
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        match_question = pattern.match(texte)
        if match_question:
            q_num = match_question.group(1)
            correct = corrections.get(q_num, '?')
            current_question = {
                "numero": q_num,
                "index": i,
                "texte": f"{q_num} - {match_question.group(2)}?",
                "reponses": [],
                "correct_lettre": correct
            }
            questions.append(current_question)
        
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1)
                texte_rep = match_reponse.group(2).strip()
                is_correct = lettre == current_question["correct_lettre"]
                
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "correct": is_correct
                })
    
    return [q for q in questions if q["reponses"]]

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file and correction_file:
    if 'questions' not in st.session_state:
        corrections = charger_corrections(correction_file)
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc, corrections)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}

    st.markdown("### Configuration des questions")
    
    for q in st.session_state.questions:
        q_id = q['index']
        q_num = q['numero']
        
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
                default_idx = next((i for i, r in enumerate(q['reponses']) if r['correct'] else 0, 0)
                
                bonne = st.selectbox(
                    f"Bonne réponse pour {q_num}",
                    options=options,
                    index=default_idx,
                    key=f"bonne_{q_id}"
                )
                
                st.session_state.figees[q_id] = True
                st.session_state.reponses_correctes[q_id] = options.index(bonne)

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION
# =============================================
def generer_document(row, template_path):
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        
        # Calcul des résultats
        modules = defaultdict(int)
        total_correct = 0
        
        for q in st.session_state.questions:
            module = q['numero'].split('.')[0]
            if st.session_state.figees.get(q['index'], False):
                if st.session_state.reponses_correctes.get(q['index'], 0) == 0:
                    modules[module] += 1
                    total_correct += 1
            else:
                if any(r['correct'] for r in q['reponses']):
                    modules[module] += 1
                    total_correct += 1
        
        # Ajout des résultats
        for mod in range(1, 6):
            replacements[f'{{result_mod{mod}}}'] = str(modules.get(str(mod), 0))
        replacements['{{result_mod_total}}'] = str(total_correct)
        
        # Remplacement des variables
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, value)
        
        # Traitement des questions
        for q in st.session_state.questions:
            reponses = q['reponses'].copy()
            is_figee = st.session_state.figees.get(q['index'], False)
            
            if is_figee:
                # Utiliser la réponse figée
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], 0)
                reponse_correcte = reponses.pop(bonne_idx)
                reponses.insert(0, reponse_correcte)
            else:
                random.shuffle(reponses)
                correct_idx = next((i for i, r in enumerate(reponses) if r['correct']), None)
                if correct_idx is not None:
                    reponse_correcte = reponses.pop(correct_idx)
                    reponses.insert(0, reponse_correcte)
            
            # Mise à jour des réponses
            for i, rep in enumerate(reponses):
                para = doc.paragraphs[rep['index']]
                checkbox = "☑" if i == 0 else "☐"
                para.text = f"{rep['lettre']} - {rep['texte']} {checkbox}"
        
        return doc
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

# =============================================
# SECTION 5: GÉNÉRATION PRINCIPALE
# =============================================
if excel_file and word_file and correction_file and st.session_state.get('questions'):
    if st.button("Générer les QCM", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                df = pd.read_excel(excel_file)
                required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
                
                if not all(col in df.columns for col in required_cols):
                    missing = [col for col in required_cols if col not in df.columns]
                    st.error(f"Colonnes manquantes : {', '.join(missing)}")
                    st.stop()

                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    f.write(word_file.getbuffer())

                zip_path = os.path.join(tmpdir, "QCM_Generes.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    
                    for idx, row in df.iterrows():
                        try:
                            doc = generer_document(row, template_path)
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            doc.save(os.path.join(tmpdir, filename))
                            zipf.write(os.path.join(tmpdir, filename), filename)
                            progress_bar.progress((idx + 1) / len(df))
                        except Exception as e:
                            st.error(f"Échec pour {row['Prénom']} {row['Nom']} : {str(e)}")
                            continue

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
