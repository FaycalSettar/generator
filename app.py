import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés avec contrôle des questions")

# Upload des fichiers
excel_file = st.file_uploader("1. Importer le fichier Excel (colonnes : Prénom, Nom)", type="xlsx")
word_file = st.file_uploader("2. Importer le modèle Word (avec {{prenom}} et {{nom}})", type="docx")

# Fonction : extraire les questions du fichier Word
def extraire_questions(doc):
    questions = []
    i = 0
    while i < len(doc.paragraphs):
        text = doc.paragraphs[i].text.strip()
        if text.endswith('?') and text[0].isdigit():
            question = {'index': i, 'texte': text}
            questions.append(question)
        i += 1
    return questions

questions_figees = []

if word_file:
    doc_temp = Document(word_file)
    questions = extraire_questions(doc_temp)
    st.markdown("### 3. Choisissez les questions à figer (ne pas mélanger les réponses)")
    for q in questions:
        key = f"figer_{q['index']}"
        if st.checkbox(f"Figer : {q['texte']}", key=key):
            questions_figees.append(q['index'])

# Fonction : mélanger les réponses après une question
def melanger_reponses(paragraphs, index_question):
    reponses = []
    i = index_question + 1
    while i < len(paragraphs) and paragraphs[i].text.strip().startswith(tuple("ABCD")):
        reponses.append(paragraphs[i].text)
        i += 1
    reponses_melangees = random.sample(reponses, len(reponses))
    for j in range(len(reponses)):
        paragraphs[index_question + 1 + j].text = reponses_melangees[j]

# Zones d’affichage en direct
log_zone = st.empty()
progress_bar = st.progress(0)
percent_display = st.empty()

# Bouton de génération
if st.button("4. Générer les fichiers QCM") and excel_file and word_file:
    with tempfile.TemporaryDirectory() as tmpdirname:
        try:
            df = pd.read_excel(excel_file)
            word_path = os.path.join(tmpdirname, "template.docx")
            with open(word_path, "wb") as f:
                f.write(word_file.read())

            zip_path = os.path.join(tmpdirname, "QCM_generes.zip")
            with ZipFile(zip_path, 'w') as zipf:
                total = len(df)

                for i, row in df.iterrows():
                    prenom, nom = row["Prénom"], row["Nom"]
                    doc = Document(word_path)

                    # Remplacer les tags
                    for para in doc.paragraphs:
                        para.text = para.text.replace("{{prenom}}", str(prenom)).replace("{{nom}}", str(nom))

                    # Mélanger les réponses sauf pour les questions figées
                    j = 0
                    while j < len(doc.paragraphs):
                        if doc.paragraphs[j].text.strip().endswith("?") and j not in questions_figees:
                            melanger_reponses(doc.paragraphs, j)
                        j += 1

                    filename = f"QCM_{prenom}_{nom}.docx"
                    filepath = os.path.join(tmpdirname, filename)
                    doc.save(filepath)
                    zipf.write(filepath, arcname=filename)

                    progress = (i + 1) / total
                    progress_bar.progress(progress)
                    percent_display.write(f"Progression : {int(progress * 100)} %")
                    log_zone.write(f"✅ Fichier généré : {filename}")

            with open(zip_path, "rb") as f:
                st.success("Tous les fichiers ont été générés avec succès !")
                st.download_button("Télécharger le ZIP contenant tous les QCM", f, file_name="QCM_personnalises.zip")

        except Exception as e:
            st.error(f"Erreur : {str(e)}")
app.py
Affichage de app.py en cours...
