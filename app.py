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
    word_file = st.file_uploader("Modèle Word (doit contenir {{checkbox}})", type="docx")

# =============================================
# SECTION 2: CONFIGURATION DES QUESTIONS
# =============================================
def detecter_questions(doc):
    """Détection améliorée des questions dans le document Word"""
    questions = []
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Vérification du pattern: numéro + checkbox + texte + ?
        if re.match(r'^\d+\.\s*{{checkbox}}.*\?$', texte):
            reponses = []
            j = i + 1
            while j < len(doc.paragraphs):
                ligne = doc.paragraphs[j].text.strip()
                if re.match(r'^[A-D]\.\s+.+', ligne):
                    reponses.append(ligne)
                    j += 1
                else:
                    break
            if len(reponses) >= 2:
                questions.append({
                    "index": i,
                    "texte": texte,
                    "reponses": reponses
                })
    return questions

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
            
            for q in questions:
                col1, col2 = st.columns([1, 5])
                with col1:
                    figer = st.checkbox(
                        f"Q{q['texte'].split('.')[0]}",
                        key=f"figer_{q['index']}",
                        help="Cocher pour figer cette question"
                    )
                with col2:
                    if figer:
                        options = [r.split(' ', 1)[1] for r in q['reponses']]
                        bonne = st.selectbox(
                            f"Bonne réponse pour Q{q['texte'].split('.')[0]}",
                            options=options,
                            key=f"bonne_{q['index']}",
                            format_func=lambda x: x
                        )
                        st.session_state.figees[q["index"]] = True
                        st.session_state.reponses_correctes[q["index"]] = f"{q['reponses'][options.index(bonne)][0]} {bonne}"
        
        st.session_state.config_done = True

    except Exception as e:
        st.error(f"Erreur de lecture du fichier Word: {str(e)}")
        st.stop()

# =============================================
# SECTION 3: GENERATION DES QCM
# =============================================
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

        # Remplacement des variables générales
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, value)

        # Traitement des questions
        for q in questions:
            idx = q['index']
            
            # Remplacement checkbox
            checkbox = "☑" if idx in st.session_state.figees else "☐"
            doc.paragraphs[idx].text = doc.paragraphs[idx].text.replace("{{checkbox}}", checkbox)
            
            # Gestion des réponses
            if idx in st.session_state.figees:
                bonne = st.session_state.reponses_correctes[idx]
                for key, value in replacements.items():
                    bonne = bonne.replace(key, value)
                
                # Réordonnancement
                reponses = [r.text for r in doc.paragraphs[idx+1:idx+1+len(q['reponses'])]]
                if bonne in reponses:
                    reponses.remove(bonne)
                    reponses = [bonne] + reponses
                else:
                    st.error(f"Réponse correcte non trouvée pour Q{q['texte'].split('.')[0]}")
                
                for i, r in enumerate(reponses):
                    doc.paragraphs[idx+1+i].text = r
            else:
                # Mélange aléatoire
                reponses = [r.text for r in doc.paragraphs[idx+1:idx+1+len(q['reponses'])]]
                random.shuffle(reponses)
                for i, r in enumerate(reponses):
                    doc.paragraphs[idx+1+i].text = r

        return doc
    except Exception as e:
        st.error(f"Erreur lors de la génération pour {row['Prénom']} {row['Nom']}: {str(e)}")
        raise

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
