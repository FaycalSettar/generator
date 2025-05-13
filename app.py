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
    """Détection ultra-flexible des questions et réponses"""
    questions = []
    current_question = None
    patterns = {
        'question': r'^\d+\.\d+[\s\-\–\—\)\.]*.*\?$',
        'reponse': r'^[A-D][\s\)\-\–\—\.]+.*{{checkbox}}'
    }
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions
        if re.match(patterns['question'], texte, re.IGNORECASE):
            current_question = {
                'index': i,
                'texte': re.sub(r'\s+', ' ', texte),  # Nettoyage des espaces multiples
                'reponses': []
            }
            questions.append(current_question)
        
        # Détection des réponses
        elif current_question and re.match(patterns['reponse'], texte, re.IGNORECASE):
            # Extraction de la lettre et nettoyage
            lettre = texte[0].upper()
            texte_clean = re.sub(r'\s*{{checkbox}}\s*', '', texte[1:]).strip()
            texte_clean = re.sub(r'^[\-\–\—\)\.\s]+', '', texte_clean)  # Nettoyage des séparateurs
            
            current_question['reponses'].append({
                'index': i,
                'lettre': lettre,
                'texte': texte_clean
            })
    
    # Validation finale des questions
    questions_valides = []
    for q in questions:
        if len(q['reponses']) >= 2:
            questions_valides.append(q)
        else:
            st.error(f"Question ignorée : {q['texte']} ({len(q['reponses'])} réponse(s) détectée(s))")
    
    return questions_valides

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file:
    # Initialisation de la session
    if 'questions' not in st.session_state:
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}
    
    st.markdown("### Configuration des questions")
    
    # Affichage des questions
    for q in st.session_state.questions:
        q_id = q['index']
        q_title = re.split(r'\d+\.\d+\s*', q['texte'])[-1][:50] + '...'
        
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(
                f"Q{q['texte'].split()[0]}",
                value=st.session_state.figees.get(q_id, False),
                key=f"figer_{q_id}",
                help=q['texte']
            )
        
        with col2:
            if figer:
                options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                
                if not options:
                    st.error("Aucune réponse valide détectée !")
                    continue
                
                default_idx = st.session_state.reponses_correctes.get(q_id, 0)
                bonne = st.selectbox(
                    f"Bonne réponse pour {q['texte'].split()[0]}",
                    options=options,
                    index=default_idx,
                    key=f"bonne_{q_id}"
                )
                
                # Validation et stockage
                if bonne in options:
                    st.session_state.figees[q_id] = True
                    st.session_state.reponses_correctes[q_id] = options.index(bonne)
                else:
                    st.error("Erreur de sélection - veuillez recharger le template")
                    st.stop()

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION
# =============================================
def generer_document(row, template_path):
    """Génération robuste avec vérifications"""
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
                
            # Gestion des réponses
            reponses = q['reponses'].copy()
            if st.session_state.figees.get(q['index'], False):
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], 0)
                if 0 <= bonne_idx < len(reponses):
                    bonne_reponse = reponses.pop(bonne_idx)
                    reponses.insert(0, bonne_reponse)
            else:
                random.shuffle(reponses)

            # Mise à jour du document
            for i, rep in enumerate(reponses):
                para = doc.paragraphs[rep['index']]
                checkbox = "☑" if (i == 0 and st.session_state.figees.get(q['index'])) else "☐"
                para.text = f"{rep['lettre']} - {rep['texte']} {checkbox}"

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

                # Sauvegarde du template
                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    f.write(word_file.getbuffer())

                # Création de l'archive
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

                # Téléchargement final
                with open(zip_path, "rb") as f:
                    st.success("✅ Génération terminée avec succès !")
                    st.download_button(
                        "📥 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Personnalises.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"ERREUR : {str(e)}")
                st.text(traceback.format_exc())
