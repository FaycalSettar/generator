import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés avec contrôle avancé")

# Configuration des uploaders de fichiers
excel_file = st.file_uploader("1. Importer le fichier Excel (colonnes obligatoires : Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
word_file = st.file_uploader("2. Importer le modèle Word (avec variables {{checkbox}}, {{prenom}}, {{nom}}, etc.)", type="docx")

def extraire_questions_et_reponses(doc):
    """Extrait les questions et réponses du template Word"""
    questions = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text.startswith(tuple(str(n) for n in range(10))) and '{{checkbox}}' in text and text.endswith('?'):
            reponses = []
            j = i + 1
            while j < len(doc.paragraphs) and doc.paragraphs[j].text.strip().startswith(tuple("ABCD")):
                reponses.append(doc.paragraphs[j].text.strip())
                j += 1
            questions.append({
                "index": i,
                "texte": text,
                "reponses": reponses
            })
    return questions

# Initialisation des variables de session
if 'figees' not in st.session_state:
    st.session_state.figees = {}
if 'reponses_correctes' not in st.session_state:
    st.session_state.reponses_correctes = {}

# Interface de configuration des questions
if word_file:
    doc_temp = Document(word_file)
    questions_data = extraire_questions_et_reponses(doc_temp)
    
    st.markdown("### 3. Configuration des questions")
    st.write("Sélectionnez les questions à figer et choisissez la bonne réponse :")
    
    for q in questions_data:
        col1, col2 = st.columns([2, 4])
        with col1:
            figer = st.checkbox(
                f"Question {q['texte'].split('.')[0]}",
                key=f"figer_{q['index']}",
                help=q['texte']
            )
        with col2:
            if figer:
                reponses = [r.split(' ', 1)[1] for r in q["reponses"]]
                bonne = st.selectbox(
                    f"Bonne réponse pour : {q['texte'][:50]}...",
                    options=reponses,
                    key=f"bonne_{q['index']}",
                    format_func=lambda x: x
                )
                st.session_state.figees[q["index"]] = True
                st.session_state.reponses_correctes[q["index"]] = f"{q['reponses'][reponses.index(bonne)][0]} {bonne}"
            else:
                if q["index"] in st.session_state.figees:
                    del st.session_state.figees[q["index"]]
                if q["index"] in st.session_state.reponses_correctes:
                    del st.session_state.reponses_correctes[q["index"]]

def melanger_reponses(paragraphs, index_question):
    """Mélange aléatoirement les réponses d'une question"""
    reponses = []
    i = index_question + 1
    while i < len(paragraphs) and paragraphs[i].text.strip().startswith(tuple("ABCD")):
        reponses.append(paragraphs[i].text.strip())
        i += 1
    random.shuffle(reponses)
    for j in range(len(reponses)):
        paragraphs[index_question + 1 + j].text = reponses[j]

def figer_reponses(paragraphs, index_question, bonne_reponse):
    """Ordonne les réponses avec la bonne réponse en première position"""
    reponses = []
    i = index_question + 1
    while i < len(paragraphs) and paragraphs[i].text.strip().startswith(tuple("ABCD")):
        reponses.append(paragraphs[i].text.strip())
        i += 1
    
    if bonne_reponse in reponses:
        reponses.remove(bonne_reponse)
        reponses_ordonnees = [bonne_reponse] + reponses
    else:
        reponses_ordonnees = reponses
    
    for j in range(len(reponses_ordonnees)):
        paragraphs[index_question + 1 + j].text = reponses_ordonnees[j]

# Section de génération des fichiers
if st.button("4. Générer les QCM personnalisés") and excel_file and word_file:
    with tempfile.TemporaryDirectory() as tmpdirname:
        try:
            # Vérification des colonnes Excel
            df = pd.read_excel(excel_file)
            required_columns = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
            
            if not all(col in df.columns for col in required_columns):
                missing = [col for col in required_columns if col not in df.columns]
                raise ValueError(f"Colonnes manquantes dans l'Excel : {', '.join(missing)}")
            
            # Préparation du template
            word_path = os.path.join(tmpdirname, "template.docx")
            with open(word_path, "wb") as f:
                f.write(word_file.getbuffer())

            # Création de l'archive ZIP
            zip_path = os.path.join(tmpdirname, "QCM_generes.zip")
            with ZipFile(zip_path, 'w') as zipf:
                total = len(df)
                progress_bar = st.progress(0)
                status_text = st.empty()

                for index, row in df.iterrows():
                    # Extraction des données
                    prenom = str(row['Prénom']).strip()
                    nom = str(row['Nom']).strip()
                    email = str(row['Email']).strip()
                    ref_session = str(row['Référence Session']).strip()
                    date_eval = str(row['Date Évaluation']).strip()

                    # Nettoyage des noms de fichier
                    safe_prenom = re.sub(r'[\\/*?:"<>|]', '_', prenom)
                    safe_nom = re.sub(r'[\\/*?:"<>|]', '_', nom)

                    # Création du document
                    doc = Document(word_path)
                    
                    # Dictionnaire de remplacement
                    replacements = {
                        '{{prenom}}': prenom,
                        '{{nom}}': nom,
                        '{{email}}': email,
                        '{{ref_session}}': ref_session,
                        '{{date_evaluation}}': date_eval
                    }

                    # Remplacement général des variables
                    for para in doc.paragraphs:
                        for key, value in replacements.items():
                            para.text = para.text.replace(key, value)

                    # Traitement des questions et réponses
                    for q in questions_data:
                        para_index = q['index']
                        
                        # Remplacement de la checkbox
                        checkbox_status = "☑" if para_index in st.session_state.figees else "☐"
                        doc.paragraphs[para_index].text = doc.paragraphs[para_index].text.replace("{{checkbox}}", checkbox_status)
                        
                        # Gestion des réponses
                        if para_index in st.session_state.figees:
                            bonne_original = st.session_state.reponses_correctes[para_index]
                            bonne_replaced = bonne_original
                            for key, value in replacements.items():
                                bonne_replaced = bonne_replaced.replace(key, value)
                            figer_reponses(doc.paragraphs, para_index, bonne_replaced)
                        else:
                            melanger_reponses(doc.paragraphs, para_index)

                    # Sauvegarde du fichier
                    filename = f"QCM_{safe_prenom}_{safe_nom}_{ref_session}.docx"
                    filepath = os.path.join(tmpdirname, filename)
                    doc.save(filepath)
                    zipf.write(filepath, arcname=filename)

                    # Mise à jour de la progression
                    progress = (index + 1) / total
                    progress_bar.progress(progress)
                    status_text.write(f"Progression : {int(progress*100)}% - {index+1}/{total} fichiers générés")

            # Téléchargement final
            with open(zip_path, "rb") as f:
                st.success("🎉 Génération terminée avec succès !")
                st.download_button(
                    label="📥 Télécharger l'archive ZIP",
                    data=f,
                    file_name="QCM_personnalises.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"🚨 Erreur lors de la génération : {str(e)}")
            st.error("Vérifiez : format des fichiers, noms de colonnes, et données manquantes.")
