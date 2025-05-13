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
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")

# =============================================
# SECTION 2: DÉTECTION DES QUESTIONS (VERSION FINALE)
# =============================================
def detecter_questions(doc):
    """Détection précise des questions et réponses avec regex améliorée"""
    questions = []
    current_question = None
    sep_pattern = r'[\s\-–—).]+'
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions
        if re.match(r'^\d+\.\d+[\.\s\-–—]*.*\?$', texte):
            current_question = {
                "index": i,
                "texte": texte,
                "reponses": []
            }
            questions.append(current_question)
        
        # Détection des réponses avec séparateur générique
        elif current_question and re.match(r'^[A-D][\s\-–—).]+', texte):
            match = re.match(r'^([A-D])' + sep_pattern + r'(.*)', texte)
            if match:
                lettre = match.group(1).upper()
                reponse_texte = match.group(2).strip()
                
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": reponse_texte
                })
    
    # Validation finale
    return [q for q in questions if len(q['reponses']) >= 2]

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file:
    # Initialisation session
    if 'questions' not in st.session_state:
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}
    
    st.markdown("### Configuration des questions")
    
    # Affichage des questions
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
                options = [f"{r['lettre']} - {r['texte'].replace('{{checkbox}}', '').strip()" for r in q['reponses']]
                
                if not options:
                    st.error("Aucune réponse valide !")
                    continue
                
                default_idx = st.session_state.reponses_correctes.get(q_id, 0)
                bonne = st.selectbox(
                    f"Bonne réponse pour {q_num}",
                    options=options,
                    index=default_idx,
                    key=f"bonne_{q_id}"
                )
                
                st.session_state.figees[q_id] = True
                st.session_state.reponses_correctes[q_id] = options.index(bonne)

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION (CORRIGÉE)
# =============================================
def generer_document(row, template_path):
    """Génération robuste avec gestion correcte des checkboxes"""
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }

        # Remplacement des variables globales
        for para in doc.paragraphs:
            for key, value in replacements.items():
                para.text = para.text.replace(key, value)

        # Traitement des questions
        for q in st.session_state.questions:
            if not q['reponses']:
                continue
                
            reponses = q['reponses'].copy()
            is_figee = st.session_state.figees.get(q['index'], False)
            
            if is_figee:
                # Déplacer la bonne réponse en première position
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], 0)
                if 0 <= bonne_idx < len(reponses):
                    reponse_correcte = reponses.pop(bonne_idx)
                    reponses.insert(0, reponse_correcte)
            
            else:
                random.shuffle(reponses)

            # Mise à jour des réponses
            for i, rep in enumerate(reponses):
                para = doc.paragraphs[rep['index']]
                checkbox = "☑" if (i == 0 and is_figee) else "☐"
                
                # Construction du texte avec checkbox
                texte_original = rep['texte'].replace('{{checkbox}}', checkbox)
                para.text = f"{rep['lettre']} - {texte_original}"

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
