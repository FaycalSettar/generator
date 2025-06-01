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

def remplacer_placeholders(paragraph, replacements):
    """Remplace les placeholders en préservant la mise en forme avec gestion des espaces"""
    if not paragraph.text:
        return
    
    # Nettoyage et préparation du texte
    original_text = paragraph.text
    
    # Parcours de tous les remplacements à effectuer
    for key, value in replacements.items():
        # Création de variantes de la clé pour gérer les espaces
        key_variants = [
            key,  # {{checkbox}}
            key.replace(" ", " "),  # {{ checkbox }}
            key.replace(" ", ""),  # {{checkbox}} sans espaces
            key.replace(" ", "").replace("{", "").replace("}", "")  # checkbox seul
        ]
        
        # Recherche et remplacement de toutes les variantes
        for variant in key_variants:
            if variant in original_text:
                # Nettoyage complet du paragraphe
                paragraph.clear()
                # Reconstruction du texte avec le remplacement
                new_text = original_text.replace(variant, value)
                paragraph.add_run(new_text)
                break

def detecter_questions(doc):
    """Détecte les questions et modules avec gestion des formats complexes"""
    questions = []
    current_question = None
    
    # Regex améliorée pour gérer les formats spécifiques du fichier
    question_pattern = re.compile(r'^(?:Module\s*\d+|[1-9]\d*(?:\.\d+)?)\s*[-–—)\s.]*\s*(.+?)(\?|$)')
    reponse_pattern = re.compile(r'^([A-D])[\s\-–—).]+\s*(.*?)\s*({{ ?checkbox ?}})?\s*$')
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions
        match_question = question_pattern.match(texte)
        if match_question:
            # Extraction du module (si présent)
            module_match = re.match(r'(Module\s*\d+)', texte)
            if module_match:
                module = module_match.group(1)
                current_question = {
                    "index": i,
                    "texte": texte,
                    "module": module,
                    "reponses": [],
                    "correct_idx": None,
                    "original_text": texte
                }
            else:
                # Extraction du numéro de question (avec ou sans points)
                question_num_match = re.match(r'((?:\d+[\. ]*)+\d*)', texte)
                if question_num_match:
                    question_num = re.sub(r'\s+', '.', question_num_match.group(1)).strip()
                    question_num = re.sub(r'\.+$', '', question_num)  # Supprime les points en trop
                    
                    # Détection automatique du module (ex: 1.1 → Module 1)
                    module_num = question_num.split('.')[0] if '.' in question_num else "1"
                    module = f"Module {module_num}"
                    
                    current_question = {
                        "index": i,
                        "texte": texte,
                        "module": module,
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
                is_correct = bool(match_reponse.group(3))
                
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": texte_rep,
                    "correct": is_correct,
                    "original_text": texte
                })
                
                if is_correct:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1

    # Validation qu'au moins 2 réponses sont présentes et qu'une réponse correcte est définie
    return [q for q in questions if len(q["reponses"]) >= 2 and q["correct_idx"] is not None]

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

# Interface utilisateur
with st.expander("Étape 1: Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word", type="docx")
    correct_answers_file = st.file_uploader("Fichier des réponses correctes (Quizz.xlsx)", type=["xlsx"])

# Chargement des questions et corrections
if word_file:
    if 'questions' not in st.session_state or st.session_state.get('current_template') != word_file.name:
        try:
            doc = Document(word_file)
            st.session_state.questions = detecter_questions(doc)
            st.session_state.figees = {}
            st.session_state.reponses_correctes = {}
            st.session_state.current_template = word_file.name
            
            # Vérification que des questions ont été détectées
            if not st.session_state.questions:
                st.warning("⚠️ Aucune question détectée. Vérifiez que :")
                st.warning("- Les questions commencent par un numéro (ex: 1.1 - ...) ou 'Module X'")
                st.warning("- Les bonnes réponses contiennent {{checkbox}}")
                st.warning("- Les questions se terminent par un point d'interrogation")
        except Exception as e:
            st.error(f"Erreur lors du chargement du document Word : {str(e)}")

if correct_answers_file:
    st.session_state.correct_answers = parse_correct_answers(correct_answers_file)
    if st.session_state.correct_answers:
        st.success(f"✅ {len(st.session_state.correct_answers)} réponses correctes chargées")

# Configuration des questions
st.markdown("### Configuration des questions")

# Affichage d'un message si aucune question n'a été détectée
if not st.session_state.get('questions'):
    st.warning("Aucune question détectée. Vérifiez que :")
    st.warning("- Les questions commencent par un numéro (ex: 1.1 - ...)")
    st.warning("- Les bonnes réponses contiennent {{checkbox}}")
    st.warning("- Les questions se terminent par un point d'interrogation")

for q in st.session_state.get('questions', []):
    q_id = q['index']
    q_num = q['texte'].split()[0]
    col1, col2 = st.columns([1, 4])
    with col1:
        figer = st.checkbox(
            f"Q{q_num}",
            value=st.session_state.figees.get(q_id, False),
            key=f"figer_{q_id}_{word_file.name[:5]}" if word_file else f"figer_{q_id}"
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

def calculer_resultat_final(total_score, total_questions):
    """Calcule le résultat final en fonction du score total"""
    if total_questions == 0:
        return "Non acquis"
        
    pourcentage = (total_score / total_questions) * 100
    
    if pourcentage >= 75:
        return "Acquis"
    elif pourcentage >= 50:
        return "En cours d’acquisition"
    else:
        return "Non acquis"

def generer_document(row, template_path):
    """Génération avec gestion dynamique des modules"""
    try:
        doc = Document(template_path)
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        
        # Détection automatique des modules à partir des questions
        modules = set()
        for q in st.session_state.questions:
            if hasattr(q, 'module'):
                modules.add(q['module'])
        
        # Initialisation des scores par module
        scores = {}
        for module in modules:
            scores[module] = 0
        scores["Total"] = 0  # Score global
        
        # Remplacement des variables dans les paragraphes
        for para in doc.paragraphs:
            remplacer_placeholders(para, replacements)
        
        # Remplacement des variables dans les tableaux
        for table in doc.tables:
            for row_table in table.rows:
                for cell in row_table.cells:
                    for para in cell.paragraphs:
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
                texte_base = rep['original_text'].split(' ', 1)[0]
                texte_reponse = rep['texte']
                ligne_complete = f"{texte_base} - {texte_reponse} {checkbox}"
                doc.paragraphs[idx].text = ligne_complete
            
            # Vérification de la réponse
            question_num = q['texte'].split(" ")[0]
            question_num_clean = re.sub(r'[^\d.]', '', question_num.split('-')[0].strip())
            
            if question_num_clean in st.session_state.correct_answers:
                correct_answer = st.session_state.correct_answers[question_num_clean].upper()
                generated_answer = reponses[0]['lettre'].upper()
                
                if generated_answer == correct_answer:
                    # Incrémentation du score du module correspondant
                    module = q.get('module', 'Module 1')
                    scores[module] = scores.get(module, 0) + 1
                    scores["Total"] += 1
        
        # Calcul du résultat final
        total_score = scores["Total"]
        total_questions = len(st.session_state.questions)
        resultat_final = calculer_resultat_final(total_score, total_questions)
        
        # Préparation des remplacements de scores
        score_replacements = {
            '{{result_mod_total}}': str(total_score),
            '{{result_evaluation}}': resultat_final
        }
        
        # Ajout des scores par module
        for i, module in enumerate(modules, 1):
            score_replacements[f'{{{{result_mod{i}}}}'] = str(scores[module])
        
        # Remplacement dans tous les paragraphes
        for para in doc.paragraphs:
            remplacer_placeholders(para, score_replacements)
        
        # Remplacement dans les tableaux
        for table in doc.tables:
            for row_table in table.rows:
                for cell in row_table.cells:
                    for para in cell.paragraphs:
                        remplacer_placeholders(para, score_replacements)
        
        return doc
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

# Génération principale
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

# Affichage des consignes
st.markdown("### Résultat final")
st.info("""
- 75% ou plus de bonnes réponses : Acquis  
- Entre 50% et 75% de bonnes réponses : En cours d’acquisition  
- Inférieur à 50% de bonnes réponses : Non acquis
""")
