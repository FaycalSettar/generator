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
            current_question["reponses"].append({
                "index": i,
                "texte": texte,
                "lettre": texte[0]
            })
    
    return questions

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file and not st.session_state.get('config_done'):
    try:
        doc = Document(word_file)
        questions = detecter_questions(doc)
        
        if not questions:
            st.error("Aucune question détectée ! Vérifiez le format de votre template.")
            st.stop()

        st.session_state.questions = questions
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}

        with st.expander("Étape 2: Configuration des questions", expanded=True):
            st.write("Cochez les questions à figer et sélectionnez les bonnes réponses :")
            
            for q in st.session_state.questions:
                col1, col2 = st.columns([1, 5])
                with col1:
                    figer = st.checkbox(
                        f"Q{q['texte'].split(' ')[0]}",
                        key=f"figer_{q['index']}",
                        help=q['texte']
                    )
                with col2:
                    if figer:
                        options = [r['texte'].split('{{checkbox}}')[0].strip() for r in q['reponses']]
                        bonne = st.selectbox(
                            f"Bonne réponse pour {q['texte'].split(' ')[0]}",
                            options=options,
                            key=f"bonne_{q['index']}",
                            format_func=lambda x: x
                        )
                        st.session_state.figees[q["index"]] = True
                        st.session_state.reponses_correctes[q["index"]] = next(
                            r['lettre'] for r in q['reponses'] 
                            if r['texte'].split('{{checkbox}}')[0].strip() == bonne
                        )
        
        st.session_state.config_done = True

    except Exception as e:
        st.error(f"Erreur de lecture du fichier Word: {str(e)}")
        st.stop()

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION
# =============================================
def remplacer_variables(paragraphs, replacements):
    """Remplacement des variables globales"""
    for para in paragraphs:
        for key, value in replacements.items():
            if key in para.text:
                para.text = para.text.replace(key, value)

def traiter_question(doc, q, replacements, bonne_reponse=None):
    """Traitement d'une question individuelle"""
    # Remplacement des variables dans la question
    for key, value in replacements.items():
        doc.paragraphs[q['index']].text = doc.paragraphs[q['index']].text.replace(key, value)
    
    # Traitement des réponses
    for reponse in q['reponses']:
        para = doc.paragraphs[reponse['index']]
        texte = para.text.replace('{{checkbox}}', '☑' if bonne_reponse == reponse['lettre'] else '☐')
        
        # Remplacement des variables dans les réponses
        for key, value in replacements.items():
            texte = texte.replace(key, value)
        
        para.text = texte

def generer_qcm(row, template_path, questions):
    """Génère un QCM individualisé"""
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
        remplacer_variables(doc.paragraphs, replacements)

        # Traitement des questions
        for q in questions:
            bonne = st.session_state.reponses_correctes.get(q['index'])
            if st.session_state.figees.get(q['index']):
                traiter_question(doc, q, replacements, bonne)
            else:
                # Mélanger les réponses
                random.shuffle(q['reponses'])
                traiter_question(doc, q, replacements)

        return doc
    except Exception as e:
        st.error(f"Erreur lors de la génération pour {row['Prénom']} {row['Nom']}: {str(e)}")
        raise

# =============================================
# SECTION 5: GÉNÉRATION PRINCIPALE
# =============================================
if excel_file and word_file and st.session_state.get('config_done'):
    if st.button("Générer les QCM", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            try:
                # Vérification Excel
                df = pd.read_excel(excel_file)
                required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
                missing = [c for c in required_cols if c not in df.columns]
                
                if missing:
                    st.error(f"Colonnes manquantes dans Excel: {', '.join(missing)}")
                    st.stop()

                # Sauvegarde template
                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    f.write(word_file.getbuffer())

                # Création ZIP
                zip_path = os.path.join(tmpdir, "QCM.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    
                    for i, row in df.iterrows():
                        try:
                            doc = generer_qcm(row, template_path, st.session_state.questions)
                            
                            # Nom de fichier sécurisé
                            nom_fichier = f"QCM_{re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))}_{re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))}.docx"
                            doc_path = os.path.join(tmpdir, nom_fichier)
                            doc.save(doc_path)
                            zipf.write(doc_path, nom_fichier)
                            
                            progress_bar.progress((i+1)/len(df))
                            
                        except Exception as e:
                            st.error(f"Échec pour {row['Prénom']} {row['Nom']}")
                            continue

                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.success("Génération terminée !")
                    st.download_button(
                        "📥 Télécharger les QCM",
                        data=f,
                        file_name="QCM_Generes.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"ERREUR CRITIQUE: {str(e)}")
                st.text(traceback.format_exc())
