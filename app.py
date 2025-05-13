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
# SECTION 2: DÉTECTION DES QUESTIONS
# =============================================
def detecter_questions(doc):
    """Détection adaptée au template spécifique"""
    questions = []
    current_question = None
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions (format 1.1 - ... ?)
        if re.match(r'^\d+\.\d+\s*-\s*.+\?$', texte):
            current_question = {
                "index": i,
                "texte": texte,
                "reponses": []
            }
            questions.append(current_question)
        # Détection des réponses (A - ... {{checkbox}})
        elif current_question and re.match(r'^[A-D]\s*-\s*.+{{checkbox}}', texte):
            cleaned = re.sub(r'\s*{{checkbox}}\s*', '', texte)
            current_question["reponses"].append({
                "index": i,
                "texte": cleaned,
                "lettre": texte[0]
            })
    
    return questions

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file:
    if 'questions' not in st.session_state:
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}

    st.markdown("### Configuration des questions")
    st.write("Sélectionnez les questions à figer et choisissez la bonne réponse :")
    
    for q in st.session_state.questions:
        unique_key = f"q_{q['index']}"
        
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(
                f"Q{q['texte'].split(' ')[0]}",
                value=st.session_state.figees.get(q['index'], False),
                key=f"figer_{unique_key}",
                help=q['texte']
            )
        
        with col2:
            if figer:
                options = [r['texte'] for r in q['reponses']]
                default_index = st.session_state.reponses_correctes.get(q['index'], 0)
                
                bonne = st.selectbox(
                    f"Bonne réponse pour {q['texte'].split(' ')[0]}",
                    options=options,
                    index=default_index,
                    key=f"bonne_{unique_key}"
                )
                
                st.session_state.figees[q['index']] = True
                st.session_state.reponses_correctes[q['index']] = options.index(bonne)
            else:
                st.session_state.figees[q['index']] = False

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION
# =============================================
def generer_document(row, template_path):
    """Génère un document individuel"""
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }

        # Remplacement des variables générales
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, value)

        # Traitement des questions
        for q in st.session_state.questions:
            # Trouver la bonne réponse
            bonne_index = st.session_state.reponses_correctes.get(q['index'])
            reponses = q['reponses'].copy()
            
            if st.session_state.figees.get(q['index']):
                # Déplacer la bonne réponse en première position
                bonne_reponse = reponses.pop(bonne_index)
                reponses = [bonne_reponse] + reponses
            else:
                random.shuffle(reponses)

            # Mise à jour des réponses dans le document
            for i, rep in enumerate(reponses):
                para = doc.paragraphs[rep['index']]
                checkbox = "☑" if (i == 0 and st.session_state.figees.get(q['index'])) else "☐"
                para.text = f"{rep['lettre']} - {rep['texte']} {checkbox}"

        return doc
    except Exception as e:
        st.error(f"Erreur génération pour {row['Prénom']} {row['Nom']}: {str(e)}")
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

                # Création ZIP
                zip_path = os.path.join(tmpdir, "QCM.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress = st.progress(0)
                    
                    for idx, row in df.iterrows():
                        try:
                            doc = generer_document(row, template_path)
                            
                            # Nom de fichier sécurisé
                            nom_fichier = f"QCM_{re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))}_{re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))}.docx"
                            doc_path = os.path.join(tmpdir, nom_fichier)
                            doc.save(doc_path)
                            zipf.write(doc_path, nom_fichier)
                            
                            progress.progress((idx + 1) / len(df))
                            
                        except Exception as e:
                            st.error(f"Échec pour {row['Prénom']} {row['Nom']}")
                            continue

                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.success("Génération terminée avec succès !")
                    st.download_button(
                        "📥 Télécharger l'archive",
                        data=f,
                        file_name="QCM_Generes.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"ERREUR : {str(e)}")
                st.text(traceback.format_exc())
