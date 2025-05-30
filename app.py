import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re
import io
import traceback

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# Fonctions utilitaires
def remplacer_placeholders(paragraph, replacements):
    if not paragraph.text:
        return
    for key, value in replacements.items():
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)
            # Gestion des espaces insécables
            if key.replace(" ", "\u00a0") in run.text:
                run.text = run.text.replace(key.replace(" ", "\u00a0"), value)
            # Gestion sans espaces
            if key.replace(" ", "") in run.text:
                run.text = run.text.replace(key.replace(" ", ""), value)

def detecter_questions(doc):
    questions = []
    current_question = None
    question_pattern = re.compile(r'^\s*(\d+(?:\.\d+)?)\s*[-\u2013\u2014)\s.]*\s*(.+?)$')
    reponse_pattern = re.compile(r'^([A-D])\s*[-\u2013\u2014)\s.]+\s*(.*?)(\{\{checkbox\}\})?$', re.IGNORECASE)
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions (plus flexible)
        if re.search(r'^\d+(\.\d+)?\s*[-–—)\s.]', texte) and not re.search(r'^\d+\s*\/\s*\d+', texte):
            match = question_pattern.match(texte)
            if match:
                question_num = match.group(1).strip()
                question_text = match.group(2).strip()
                
                # Nettoyer le texte de la question
                if not question_text.endswith('?'):
                    question_text += '?'
                
                current_question = {
                    "index": i,
                    "texte": f"{question_num} - {question_text}",
                    "reponses": [],
                    "correct_idx": None,
                    "original_text": texte
                }
                questions.append(current_question)
                continue
        
        # Détection des réponses
        if current_question:
            match = reponse_pattern.match(texte)
            if match:
                lettre = match.group(1).upper()
                texte_rep = match.group(2).strip()
                is_correct = bool(match.group(3))
                
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "correct": is_correct,
                    "original_text": texte
                })
                
                if is_correct:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1
    
    # Filtrer les questions valides
    valid_questions = []
    for q in questions:
        if q.get('correct_idx') is not None and len(q['reponses']) >= 2:
            valid_questions.append(q)
        else:
            st.warning(f"Question ignorée: {q['texte']} - Réponse correcte non détectée ou nombre de réponses insuffisant")
    
    return valid_questions

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

def calculer_resultat_final(total_score, total_questions=9):
    pourcentage = (total_score / total_questions) * 100
    if pourcentage >= 75:
        return "Acquis"
    elif pourcentage >= 50:
        return "En cours d'acquisition"
    else:
        return "Non acquis"

def generer_document(row, doc_template):
    try:
        doc = doc_template
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        
        # Remplacer les placeholders dans tout le document
        for para in doc.paragraphs:
            remplacer_placeholders(para, replacements)
            
        for table in doc.tables:
            for row_table in table.rows:
                for cell in row_table.cells:
                    for para in cell.paragraphs:
                        remplacer_placeholders(para, replacements)
        
        # Traiter les questions
        correct_answers = st.session_state.get('correct_answers', {})
        score_total = 0
        
        for q in st.session_state.questions:
            reponses = q['reponses'].copy()
            q_num = q['texte'].split()[0].replace(':', '').strip()
            
            # Déterminer l'ordre des réponses
            is_figee = st.session_state.figees.get(q['index'], False)
            if is_figee:
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
                reponse_correcte = reponses.pop(bonne_idx)
                reponses.insert(0, reponse_correcte)
            else:
                if q['correct_idx'] is not None:
                    reponse_correcte = reponses.pop(q['correct_idx'])
                    reponses.insert(0, reponse_correcte)
                random.shuffle(reponses)
            
            # Mettre à jour le document avec les réponses
            for rep in reponses:
                para_idx = rep['index']
                if para_idx < len(doc.paragraphs):
                    checkbox = "☑" if reponses.index(rep) == 0 else "☐"
                    # Préserver la mise en forme originale
                    doc.paragraphs[para_idx].text = ""
                    run = doc.paragraphs[para_idx].add_run(f"{rep['lettre']} - {rep['texte']} {checkbox}")
            
            # Vérifier la réponse correcte
            if q_num in correct_answers:
                correct_answer = correct_answers[q_num].upper()
                generated_answer = reponses[0]['lettre'].upper()
                if generated_answer == correct_answer:
                    score_total += 1
        
        # Calculer le résultat final
        resultat_final = calculer_resultat_final(score_total)
        
        # Remplacer les résultats dans le document
        score_replacements = {
            '{{result_mod1}}': str(score_total),
            '{{result_mod_total}}': str(score_total),
            '{{result_evaluation}}': resultat_final
        }
        
        for para in doc.paragraphs:
            remplacer_placeholders(para, score_replacements)
            
        for table in doc.tables:
            for row_table in table.rows:
                for cell in row_table.cells:
                    for para in cell.paragraphs:
                        remplacer_placeholders(para, score_replacements)
        
        return doc, score_total, resultat_final
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        st.error(traceback.format_exc())
        return None, 0, "Erreur"

# Interface Streamlit
with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    correct_answers_file = st.file_uploader("Fichier des réponses correctes (Quizz.xlsx)", type=["xlsx"])

# Initialiser l'état de session
if 'questions' not in st.session_state:
    st.session_state.questions = []
if 'figees' not in st.session_state:
    st.session_state.figees = {}
if 'reponses_correctes' not in st.session_state:
    st.session_state.reponses_correctes = {}
if 'current_template' not in st.session_state:
    st.session_state.current_template = None
if 'doc_template' not in st.session_state:
    st.session_state.doc_template = None

# Charger le template Word et détecter les questions
if word_file:
    if st.session_state.get('current_template') != word_file.name:
        try:
            # Charger le document
            doc_bytes = word_file.getvalue()
            doc = Document(io.BytesIO(doc_bytes))
            st.session_state.doc_template = doc
            
            # Détecter les questions
            questions = detecter_questions(doc)
            st.session_state.questions = questions
            st.session_state.current_template = word_file.name
            st.success(f"✅ {len(questions)} questions détectées dans le document")
            
            # Afficher les questions détectées pour vérification
            with st.expander("Questions détectées (Vérification)"):
                if not questions:
                    st.warning("Aucune question détectée. Vérifiez le format de votre document.")
                for i, q in enumerate(questions):
                    st.subheader(f"Question {i+1}: {q['texte']}")
                    for j, r in enumerate(q['reponses']):
                        prefix = "✅" if j == q['correct_idx'] else "☐"
                        st.write(f"{prefix} {r['lettre']}: {r['texte']}")
                        st.caption(f"Original: '{r['original_text']}'")
        except Exception as e:
            st.error(f"Erreur lors du chargement du document Word: {str(e)}")
            st.error(traceback.format_exc())

# Charger les réponses correctes
if correct_answers_file:
    st.session_state.correct_answers = parse_correct_answers(correct_answers_file)
    if st.session_state.correct_answers:
        st.success(f"✅ {len(st.session_state.correct_answers)} réponses correctes chargées")

# Configuration des questions
if st.session_state.questions:
    st.markdown("### Configuration des questions")
    
    for q in st.session_state.questions:
        q_id = q['index']
        q_num = q['texte'].split()[0]
        
        st.markdown(f"**Question {q_num}**")
        st.write(q['texte'])
        
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(
                "Figer cette question", 
                value=st.session_state.figees.get(q_id, False), 
                key=f"figer_{q_id}"
            )
        
        if figer:
            options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
            default_idx = q['correct_idx']
            
            bonne = st.selectbox(
                f"Sélectionnez la bonne réponse pour {q_num}", 
                options=options, 
                index=default_idx, 
                key=f"bonne_{q_id}"
            )
            
            st.session_state.figees[q_id] = True
            st.session_state.reponses_correctes[q_id] = options.index(bonne)

# Génération des documents
if excel_file and st.session_state.get('questions') and st.button("Générer les QCM", type="primary"):
    try:
        # Lire le fichier Excel
        df = pd.read_excel(excel_file)
        required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
        
        # Vérifier les colonnes requises
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            st.error(f"Colonnes manquantes dans le fichier Excel: {', '.join(missing)}")
            st.stop()
        
        # Préparer le ZIP
        zip_buffer = io.BytesIO()
        recap_data = []
        
        with ZipFile(zip_buffer, 'w') as zipf:
            progress_bar = st.progress(0)
            total_rows = len(df)
            
            for idx, row in df.iterrows():
                try:
                    # Créer une copie du template pour chaque étudiant
                    doc_copy = Document(io.BytesIO(word_file.getvalue()))
                    
                    # Générer le document personnalisé
                    doc, score, resultat = generer_document(row, doc_copy)
                    
                    if doc is None:
                        st.error(f"Échec de génération pour {row['Prénom']} {row['Nom']}")
                        continue
                    
                    # Ajouter au récapitulatif
                    recap_data.append({
                        "Prénom": row["Prénom"],
                        "Nom": row["Nom"],
                        "Email": row["Email"],
                        "Référence Session": row["Référence Session"],
                        "Score": score,
                        "Résultat": resultat
                    })
                    
                    # Sauvegarder le document
                    safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                    safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                    filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                        doc.save(tmp_file.name)
                        zipf.write(tmp_file.name, filename)
                    
                    # Mettre à jour la barre de progression
                    progress_bar.progress((idx + 1) / total_rows)
                    
                except Exception as e:
                    st.error(f"Échec pour {row['Prénom']} {row['Nom']}: {str(e)}")
                    continue
        
        # Ajouter le récapitulatif au ZIP
        if recap_data:
            df_recap = pd.DataFrame(recap_data)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as recap_file:
                df_recap.to_excel(recap_file.name, index=False)
                zipf.write(recap_file.name, "Recapitulatif_QCM.xlsx")
        
        # Télécharger le ZIP
        zip_buffer.seek(0)
        st.success("✅ Génération terminée avec succès !")
        st.download_button(
            "💾 Télécharger l'archive ZIP",
            data=zip_buffer,
            file_name="QCM_Personnalises.zip",
            mime="application/zip"
        )
        
    except Exception as e:
        st.error(f"ERREUR CRITIQUE : {str(e)}")
        st.error(traceback.format_exc())

# Section d'information
st.markdown("### Résultat final")
st.info("""
- **Acquis** : 75% ou plus de bonnes réponses  
- **En cours d'acquisition** : Entre 50% et 75% de bonnes réponses  
- **Non acquis** : Moins de 50% de bonnes réponses
""")
