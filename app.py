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
st.title("Générateur de QCM personnalisés avec correction")

# ========= ETAPE 1 : UPLOAD DES FICHIERS ===========
with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    st.markdown(
        "- **1.** Fichier Excel participants (doit inclure les colonnes : Prénom, Nom, Email, Référence Session, Date Évaluation, Rép1, Rép2...)\n"
        "- **2.** Modèle Word (.docx)\n"
        "- **3.** Fichier Quizz/correction (Excel avec colonnes 'Question', 'BonneRéponse')"
    )
    excel_file = st.file_uploader("Fichier Excel Participants", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    quizz_file = st.file_uploader("Fichier de correction (Quizz)", type="xlsx")

# ========= ETAPE 2 : TRAITEMENT DES FICHIERS ===========
def injecte_scores(doc, scores, total):
    mapping = {
        '{{result_mod1}}': str(scores[0]),
        '{{result_mod2}}': str(scores[1]),
        '{{result_mod3}}': str(scores[2]),
        '{{result_mod4}}': str(scores[3]),
        '{{result_mod5}}': str(scores[4]),
        '{{result_mod_total}}': str(total)
    }
    for para in doc.paragraphs:
        for key, value in mapping.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, value)
    # Remplacement aussi dans les tableaux s'il y a lieu
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in mapping.items():
                    if key in cell.text:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.text = run.text.replace(key, value)

def score_reponses(reponses_utilisateur, correction):
    scores_par_module = [0, 0, 0, 0, 0]
    total = 0
    for i in range(1, 31):
        module_idx = (i - 1) // 6
        bonne_reponse = correction.get(str(i))
        reponse = reponses_utilisateur.get(f'Rép{i}')
        if reponse == bonne_reponse:
            scores_par_module[module_idx] += 1
            total += 1
    return scores_par_module, total

def replace_fields_in_doc(doc, replacements):
    """Remplacement des variables génériques dans le document (hors résultats modules)"""
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, value)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.text = run.text.replace(key, value)

# ========== ETAPE 3 : GÉNÉRATION PRINCIPALE ===========
if excel_file and word_file and quizz_file:
    if st.button("Générer les QCM personnalisés", type="primary"):
        try:
            df = pd.read_excel(excel_file)
            quizz_df = pd.read_excel(quizz_file)
            # Adapte le nom des colonnes ci-dessous si besoin !
            correction = dict(zip(quizz_df['Question'].astype(str), quizz_df['BonneRéponse'].astype(str)))
            # Vérifie la présence des colonnes nécessaires
            champs_participant = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
            champs_reponses = [f'Rép{i}' for i in range(1, 31)]
            for col in champs_participant + champs_reponses:
                if col not in df.columns:
                    st.error(f"Colonne manquante dans Excel : {col}")
                    st.stop()
            # Sauvegarde temporaire du modèle Word
            with tempfile.TemporaryDirectory() as tmpdir:
                template_path = os.path.join(tmpdir, "modele.docx")
                with open(template_path, "wb") as f:
                    f.write(word_file.getbuffer())
                zip_path = os.path.join(tmpdir, "QCM_Generes.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    for idx, row in df.iterrows():
                        try:
                            doc = Document(template_path)
                            # Remplace les champs classiques
                            replacements = {
                                '{{prenom}}': str(row['Prénom']),
                                '{{nom}}': str(row['Nom']),
                                '{{email}}': str(row['Email']),
                                '{{ref_session}}': str(row['Référence Session']),
                                '{{date_evaluation}}': str(row['Date Évaluation'])
                            }
                            replace_fields_in_doc(doc, replacements)
                            # Correction/scoring
                            reponses_utilisateur = {f'Rép{i}': str(row[f'Rép{i}']) for i in range(1, 31)}
                            scores, total = score_reponses(reponses_utilisateur, correction)
                            injecte_scores(doc, scores, total)
                            # Sauvegarde
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            file_path = os.path.join(tmpdir, filename)
                            doc.save(file_path)
                            zipf.write(file_path, filename)
                            progress_bar.progress((idx + 1) / len(df))
                        except Exception as e:
                            st.error(f"Erreur pour {row['Prénom']} {row['Nom']} : {str(e)}")
                            st.text(traceback.format_exc())
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
