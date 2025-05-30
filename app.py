import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# Fonctions utilitaires
def remplacer_placeholders(paragraph, replacements):
    if not paragraph.text:
        return
    original_text = paragraph.text
    for key, value in replacements.items():
        cleaned_key = key.replace(" ", "").replace("\u00a0", "")
        cleaned_text = original_text.replace(" ", "").replace("\u00a0", "")
        if cleaned_key in cleaned_text:
            for run in paragraph.runs:
                for k, v in replacements.items():
                    if k in run.text:
                        run.text = run.text.replace(k, v)
                    if k.replace(" ", "\u00a0") in run.text:
                        run.text = run.text.replace(k.replace(" ", "\u00a0"), v)
                    if k.replace(" ", "") in run.text:
                        run.text = run.text.replace(k.replace(" ", ""), v)

def detecter_questions(doc):
    questions = []
    current_question = None
    pattern = re.compile(r'^\s*(\d+(?:[., ]\d+)?)\s*[-–—)\s.]*\s*(.+?)\?$')
    reponse_pattern = re.compile(r'^([A-D])[\s\-\u2013\u2014).]+\s*(.*?)(\{\{checkbox\}\})?\s*$')
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip().replace("\u00a0", " ").replace("–", "-").replace("—", "-")
        match_question = pattern.match(texte)
        if match_question:
            question_num = re.sub(r'\s+', '.', match_question.group(1)).strip()
            question_num = re.sub(r'\.+$', '', question_num)
            current_question = {
                "index": i,
                "texte": f"{question_num} - {match_question.group(2)}?",
                "reponses": [],
                "correct_idx": None,
                "original_text": texte
            }
            questions.append(current_question)
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1)
                texte_rep = match_reponse.group(2).strip()
                is_correct = match_reponse.group(3) is not None
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "correct": is_correct,
                    "original_text": texte
                })
                if is_correct:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1
    return [q for q in questions if q["correct_idx"] is not None and len(q["reponses"]) >= 2]

def parse_correct_answers(file):
    if file is None:
        return {}
    try:
        df = pd.read_excel(file)
        df = df.dropna(subset=['Numéro de la question', 'Réponse correcte'])
        df['Numéro de la question'] = df['Numéro de la question'].astype(str).str.strip()
        df['Réponse correcte'] = df['Réponse correcte'].astype(str).str.strip().str.upper()
        return dict(zip(df['Numéro de la question'], df['Réponse correcte']))
    except Exception as e:
        st.error(f"Erreur de lecture du fichier de corrections : {str(e)}")
        return {}

with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    correct_answers_file = st.file_uploader("Fichier des réponses correctes (Quizz.xlsx)", type=["xlsx"])

if word_file:
    if 'questions' not in st.session_state or st.session_state.get('current_template') != word_file.name:
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}
        st.session_state.current_template = word_file.name

if correct_answers_file:
    st.session_state.correct_answers = parse_correct_answers(correct_answers_file)
    st.success(f"✅ {len(st.session_state.correct_answers)} réponses correctes chargées")

st.markdown("### Configuration des questions")

for q in st.session_state.get('questions', []):
    q_id = q['index']
    q_num = q['texte'].split()[0]
    col1, col2 = st.columns([1, 4])
    with col1:
        figer = st.checkbox(f"Q{q_num}", value=st.session_state.figees.get(q_id, False), key=f"figer_{q_id}_{word_file.name[:5]}" if word_file else f"figer_{q_id}")
    with col2:
        if figer:
            options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
            default_idx = q['correct_idx']
            bonne = st.selectbox(f"Bonne réponse pour {q_num}", options=options, index=default_idx, key=f"bonne_{q_id}_{word_file.name[:5]}" if word_file else f"bonne_{q_id}")
            st.session_state.figees[q_id] = True
            st.session_state.reponses_correctes[q_id] = options.index(bonne)

def calculer_resultat_final(total_score):
    total_questions = 30
    pourcentage = (total_score / total_questions) * 100
    if pourcentage >= 75:
        return "Acquis"
    elif pourcentage >= 50:
        return "En cours d’acquisition"
    else:
        return "Non acquis"

def generer_document(row, template_path):
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        scores = {f"Module {i}": 0 for i in range(1, 6)}
        correct_answers = st.session_state.get('correct_answers', {})
        for para in doc.paragraphs:
            remplacer_placeholders(para, replacements)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        remplacer_placeholders(para, replacements)
        for q in st.session_state.questions:
            reponses = q['reponses'].copy()
            is_figee = st.session_state.figees.get(q['index'], False)
            if is_figee:
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
                reponse_correcte = reponses.pop(bonne_idx)
                reponses.insert(0, reponse_correcte)
            else:
                correct_idx = next((i for i, r in enumerate(reponses) if r['correct']), None)
                if correct_idx is not None:
                    reponse_correcte = reponses.pop(correct_idx)
                    reponses.insert(0, reponse_correcte)
                random.shuffle(reponses)
            for rep in reponses:
                idx = rep['index']
                checkbox = "☑" if reponses.index(rep) == 0 else "☐"
                texte_base = rep['original_text'].split(' ', 1)[0]
                texte_reponse = rep['texte']
                ligne_complete = f"{texte_base} - {texte_reponse} {checkbox}"
                doc.paragraphs[idx].text = ligne_complete
            question_num = q['texte'].split(" ")[0]
            question_num_clean = question_num.replace(" ", ".").strip()
            question_num_clean = re.sub(r'\.+$', '', question_num_clean)
            module = f"Module {question_num_clean.split('.')[0]}"
            if question_num_clean in correct_answers:
                correct_answer = correct_answers[question_num_clean].upper()
                generated_answer = reponses[0]['lettre'].upper()
                if generated_answer == correct_answer:
                    scores[module] += 1
                    scores['Total'] = scores.get('Total', 0) + 1
        total_score = scores.get('Total', 0)
        resultat_final = calculer_resultat_final(total_score)
        score_replacements = {
            '{{result_mod1}}': str(scores["Module 1"]),
            '{{result_mod2}}': str(scores["Module 2"]),
            '{{result_mod3}}': str(scores["Module 3"]),
            '{{result_mod4}}': str(scores["Module 4"]),
            '{{result_mod5}}': str(scores["Module 5"]),
            '{{result_mod_total}}': str(total_score),
            '{{result_evaluation}}': resultat_final
        }
        for para in doc.paragraphs:
            remplacer_placeholders(para, score_replacements)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        remplacer_placeholders(para, score_replacements)
        return doc, total_score, resultat_final
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

if excel_file and word_file and st.session_state.get('questions') and st.session_state.get('correct_answers'):
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
                recapitulatif = []
                zip_path = os.path.join(tmpdir, "QCM_Generes.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    total = len(df)
                    for idx, row in df.iterrows():
                        try:
                            doc, score, resultat = generer_document(row, template_path)
                            recapitulatif.append({
                                "Prénom": row["Prénom"],
                                "Nom": row["Nom"],
                                "Référence Session": row["Référence Session"],
                                "Score /30": score,
                                "Résultat": resultat
                            })
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            doc.save(os.path.join(tmpdir, filename))
                            zipf.write(os.path.join(tmpdir, filename), filename)
                            progress_bar.progress((idx + 1)/total, text=f"Génération en cours : {idx+1}/{total}")
                        except Exception as e:
                            st.error(f"Échec pour {row['Prénom']} {row['Nom']} : {str(e)}")
                            continue
                    df_recap = pd.DataFrame(recapitulatif)
                    recap_path = os.path.join(tmpdir, "Recapitulatif_QCM.xlsx")
                    df_recap.to_excel(recap_path, index=False)
                    zipf.write(recap_path, "Recapitulatif_QCM.xlsx")
                with open(zip_path, "rb") as f:
                    st.success("✅ Génération terminée avec succès !")
                    st.download_button(
                        "💾 Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Personnalises.zip",
                        mime="application/zip"
                    )
            except Exception as e:
                st.error(f"ERREUR CRITIQUE : {str(e)}")

st.markdown("### Résultat final")
st.info("""
- 75% ou plus de bonnes réponses : Acquis  
- Entre 50% et 75% de bonnes réponses : En cours d’acquisition  
- Inférieur à 50% de bonnes réponses : Non acquis
""")
