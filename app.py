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
# FONCTIONS UTILITAIRES
# =============================================
def remplacer_placeholders(paragraph, replacements):
    """Remplace les placeholders en préservant la mise en forme"""
    for key, value in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def detecter_questions(doc):
    """Détection précise des questions avec regex améliorée"""
    questions = []
    current_question = None
    pattern = re.compile(r'^(\d+\.\d+)\s*[-–—)\s.]*\s*(.+?)\?$')
    reponse_pattern = re.compile(r'^([A-D])[\s\-–—).]+\s*(.*?)({{checkbox}})?\s*$')
   
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
       
        # Détection des questions
        match_question = pattern.match(texte)
        if match_question:
            current_question = {
                "index": i,
                "texte": f"{match_question.group(1)} - {match_question.group(2)}?",
                "reponses": [],
                "correct_idx": None,
                "original_text": texte
            }
            questions.append(current_question)
       
        # Détection des réponses
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
    """Parse le fichier Quizz.xlsx et retourne un dictionnaire {question: réponse}"""
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

# =============================================
# INTERFACE UTILISATEUR
# =============================================
with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    correct_answers_file = st.file_uploader("Fichier des réponses correctes (Quizz.xlsx)", type=["xlsx"])

# =============================================
# CHARGEMENT DES QUESTIONS ET CORRECTIONS
# =============================================
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

# =============================================
# CONFIGURATION DES QUESTIONS
# =============================================
st.markdown("### Configuration des questions")

for q in st.session_state.get('questions', []):
    q_id = q['index']
    q_num = q['texte'].split()[0]
   
    col1, col2 = st.columns([1, 4])
    with col1:
        figer = st.checkbox(
            f"Q{q_num}",
            value=st.session_state.figees.get(q_id, False),
            key=f"figer_{q_id}_{word_file.name[:5]}" if word_file else f"figer_{q_id}",
            help=q['texte']
        )
   
    with col2:
        if figer:
            options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
            default_idx = q['correct_idx']
           
            bonne = st.selectbox(
                f"Bonne réponse pour {q_num}",
                options=options,
                index=default_idx,
                key=f"bonne_{q_id}_{word_file.name[:5]}" if word_file else f"bonne_{q_id}"
            )
           
            st.session_state.figees[q_id] = True
            st.session_state.reponses_correctes[q_id] = options.index(bonne)

# =============================================
# FONCTION DE GÉNÉRATION AVEC CALCUL DES SCORES
# =============================================
def generer_document(row, template_path):
    """Génération avec calcul des scores par module"""
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }

        # Initialisation des scores
        scores = {f"Module {i}": 0 for i in range(1, 6)}
        correct_answers = st.session_state.get('correct_answers', {})
        
        # Remplacement des variables
        for para in doc.paragraphs:
            remplacer_placeholders(para, replacements)

        # Traitement des questions
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

            # Mise à jour du document avec les réponses
            for rep in reponses:
                idx = rep['index']
                checkbox = "☑" if reponses.index(rep) == 0 else "☐"
                
                # Reconstruction du texte original avec lettre et réponse
                texte_base = rep['original_text'].split(' ', 1)[0]  # Lettre de réponse
                texte_reponse = rep['texte']
                ligne_complete = f"{texte_base} - {texte_reponse} {checkbox}"
                
                doc.paragraphs[idx].text = ligne_complete

            # Vérification de la réponse
            question_num = q['texte'].split(" ")[0]  # Extrait le numéro de question (ex: "1.1")
            module = f"Module {question_num.split('.')[0]}"  # Extrait le module (ex: "Module 1")
            
            if question_num in correct_answers:
                correct_answer = correct_answers[question_num].upper()
                generated_answer = reponses[0]['lettre'].upper()
                
                if generated_answer == correct_answer:
                    scores[module] += 1
                    scores['Total'] = scores.get('Total', 0) + 1

        # Insertion des résultats dans le document
        score_replacements = {
            '{{result_mod1}}': str(scores["Module 1"]),
            '{{result_mod2}}': str(scores["Module 2"]),
            '{{result_mod3}}': str(scores["Module 3"]),
            '{{result_mod4}}': str(scores["Module 4"]),
            '{{result_mod5}}': str(scores["Module 5"]),
            '{{result_mod_total}}': str(scores.get('Total', 0))
        }

        for para in doc.paragraphs:
            remplacer_placeholders(para, score_replacements)

        return doc
        
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

# =============================================
# GÉNÉRATION PRINCIPALE
# =============================================
if excel_file and word_file and st.session_state.get('questions') and st.session_state.get('correct_answers'):
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
                    total = len(df)
                   
                    for idx, row in df.iterrows():
                        try:
                            doc = generer_document(row, template_path)
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            doc.save(os.path.join(tmpdir, filename))
                            zipf.write(os.path.join(tmpdir, filename), filename)
                            progress_bar.progress((idx + 1)/total, text=f"Génération en cours : {idx+1}/{total}")
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
