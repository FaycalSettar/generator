import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés avec contrôle avancé")

excel_file = st.file_uploader("1. Importer le fichier Excel (colonnes : Prénom, Nom)", type="xlsx")
word_file = st.file_uploader("2. Importer le modèle Word (avec {{prenom}} et {{nom}})", type="docx")

def extraire_questions_et_reponses(doc):
    questions = []
    i = 0
    while i < len(doc.paragraphs):
        question_text = doc.paragraphs[i].text.strip()
        if question_text.endswith('?') and question_text[0].isdigit():
            reponses = []
            j = i + 1
            while j < len(doc.paragraphs) and doc.paragraphs[j].text.strip().startswith(tuple("ABCD")):
                reponses.append(doc.paragraphs[j].text.strip())
                j += 1
            questions.append({
                "index": i,
                "texte": question_text,
                "reponses": reponses
            })
        i += 1
    return questions

figees = {}
reponses_correctes = {}

if word_file:
    doc_temp = Document(word_file)
    questions_data = extraire_questions_et_reponses(doc_temp)
    st.markdown("### 3. Choisissez les questions à figer et indiquez la bonne réponse à afficher en premier")
    for q in questions_data:
        col1, col2 = st.columns([2, 2])
        with col1:
            figer = st.checkbox(f"Figer : {q['texte']}", key=f"figer_{q['index']}")
            if figer:
                figees[q["index"]] = True
        with col2:
            if f"figer_{q['index']}" in st.session_state and st.session_state[f"figer_{q['index']}"]:
                bonne = st.selectbox(
                    f"Bonne réponse pour '{q['texte'][:30]}...'", 
                    options=q["reponses"],
                    key=f"bonne_{q['index']}"
                )
                reponses_correctes[q["index"]] = bonne

def ordonner_reponses_figees(bonne, reponses):
    autres = [r for r in reponses if r != bonne]
    return [bonne] + autres

def melanger_reponses(paragraphs, index_question):
    reponses = []
    i = index_question + 1
    while i < len(paragraphs) and paragraphs[i].text.strip().startswith(tuple("ABCD")):
        reponses.append(paragraphs[i].text.strip())
        i += 1
    reponses_melangees = random.sample(reponses, len(reponses))
    for j in range(len(reponses_melangees)):
        paragraphs[index_question + 1 + j].text = reponses_melangees[j]

def figer_reponses(paragraphs, index_question, bonne):
    reponses = []
    i = index_question + 1
    while i < len(paragraphs) and paragraphs[i].text.strip().startswith(tuple("ABCD")):
        reponses.append(paragraphs[i].text.strip())
        i += 1
    reponses_ordonnees = ordonner_reponses_figees(bonne, reponses)
    for j in range(len(reponses_ordonnees)):
        paragraphs[index_question + 1 + j].text = reponses_ordonnees[j]

log_zone = st.empty()
progress_bar = st.progress(0)
percent_display = st.empty()

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

                    for para in doc.paragraphs:
                        para.text = para.text.replace("{{prenom}}", str(prenom)).replace("{{nom}}", str(nom))

                    j = 0
                    while j < len(doc.paragraphs):
                        if doc.paragraphs[j].text.strip().endswith("?"):
                            if j in figees and j in reponses_correctes:
                                figer_reponses(doc.paragraphs, j, reponses_correctes[j])
                            else:
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
