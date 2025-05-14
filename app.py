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
    """Charge les corrections depuis le fichier Excel"""
    df = pd.read_excel(fichier, header=None, usecols="B,C", skiprows=1)
    df.columns = ['Question', 'Reponse']
    return {row['Question']: row['Reponse'] for _, row in df.iterrows()}

def detecter_questions(doc, corrections):
    """Détection des questions avec vérification des corrections"""
    questions = []
    current_question = None
    pattern = re.compile(r'^(\d+\.\d+)\s*[-–—)\s.]*\s*(.+?)\?$')
    reponse_pattern = re.compile(r'^([A-D])[\s\-–—).]+\s*(.*?)({{checkbox}})?\s*$')
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions
        match_question = pattern.match(texte)
        if match_question:
            q_num = match_question.group(1)
            correct = corrections.get(q_num, '?')
            current_question = {
                "numero": q_num,
                "index": i,
                "texte": f"{q_num} - {match_question.group(2)}?",
                "reponses": [],
                "correct": correct
            }
            questions.append(current_question)
        
        # Détection des réponses
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1)
                texte_rep = match_reponse.group(2).strip()
                
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "est_correcte": lettre == current_question["correct"]
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

    st.markdown("### Configuration des questions")
    
    # Calcul des résultats par module
    modules = defaultdict(int)
    for q in st.session_state.questions:
        module = q['numero'].split('.')[0]
        modules[module] += 1

    # =============================================
    # SECTION 4: FONCTIONS DE GÉNÉRATION
    # =============================================
    def generer_document(row, template_path):
        """Génération avec calcul des résultats"""
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
            resultats = defaultdict(int)
            for q in st.session_state.questions:
                module = q['numero'].split('.')[0]
                resultats[f"result_mod{module}"] += 1
            
            # Ajout des résultats aux replacements
            for mod in range(1, 6):
                replacements[f'{{result_mod{mod}}}'] = str(resultats.get(f'result_mod{mod}', 0))
            replacements['{{result_mod_total}}'] = str(sum(modules.values()))
            
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
                    # Réponses figées
                    correct_idx = next((i for i, r in enumerate(reponses) if r['est_correcte']), 0)
                    reponse_correcte = reponses.pop(correct_idx)
                    reponses.insert(0, reponse_correcte)
                else:
                    random.shuffle(reponses)
                    correct_idx = next((i for i, r in enumerate(reponses) if r['est_correcte']), None)
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
