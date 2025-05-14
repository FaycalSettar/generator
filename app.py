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
    col1, col2, col3 = st.columns(3)
    
    with col1:
        excel_file = st.file_uploader("Fichier Participants", type="xlsx", 
                                   help="Colonnes requises: Prénom, Nom, Email, Référence Session, Date Évaluation")
    
    with col2:
        word_file = st.file_uploader("Modèle Word", type="docx")
    
    with col3:
        reponses_file = st.file_uploader("Fichier Réponses", type="xlsx",
                                     help="Format attendu: colonnes 'Numéro de la question' et 'Réponse correcte'")

# =============================================
# SECTION 2: DÉTECTION DES QUESTIONS
# =============================================
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
                "module": match_question.group(1).split('.')[0] if '.' in match_question.group(1) else '1'
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
                    "correct": is_correct
                })
               
                if is_correct:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1
   
    return [q for q in questions if q["correct_idx"] is not None and len(q["reponses"]) >= 2]

# =============================================
# SECTION 3: CONFIGURATION DES QUESTIONS
# =============================================
if word_file:
    if 'questions' not in st.session_state:
        doc = Document(word_file)
        st.session_state.questions = detecter_questions(doc)
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}

    st.markdown("### Configuration des questions")
   
    # Chargement des réponses correctes depuis le fichier uploadé
    if reponses_file:
        try:
            df_reponses = pd.read_excel(reponses_file)
            
            # Vérifier que les colonnes attendues existent
            if 'Numéro de la question' in df_reponses.columns and 'Réponse correcte' in df_reponses.columns:
                st.session_state.reponses_correctes_auto = {
                    row['Numéro de la question']: row['Réponse correcte'].strip()
                    for _, row in df_reponses.iterrows()
                    if pd.notna(row['Numéro de la question']) and pd.notna(row['Réponse correcte'])
                }
            else:
                st.warning("⚠️ Le fichier de réponses doit contenir les colonnes 'Numéro de la question' et 'Réponse correcte'")
        except Exception as e:
            st.error(f"Erreur lecture réponses: {str(e)}")

    for q in st.session_state.questions:
        q_id = q['index']
        q_num = q['texte'].split()[0]
       
        col1, col2 = st.columns([1, 4])
        with col1:
            figer = st.checkbox(
                f"Q{q_num}",
                value=st.session_state.figees.get(q_id, False),
                key=f"figer_{q_id}",
                help=q['texte']
            )
       
        with col2:
            if figer:
                options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                
                # Détection automatique depuis le fichier de réponses
                default_idx = q['correct_idx'] if 'correct_idx' in q else 0
                
                if reponses_file and hasattr(st.session_state, 'reponses_correctes_auto'):
                    if q_num in st.session_state.reponses_correctes_auto:
                        auto_rep = st.session_state.reponses_correctes_auto[q_num]
                        default_idx = next((i for i, r in enumerate(q['reponses']) 
                                          if r['lettre'].strip() == auto_rep), default_idx)
                
                bonne = st.selectbox(
                    f"Bonne réponse pour {q_num} (Module {q['module']})",
                    options=options,
                    index=default_idx,
                    key=f"bonne_{q_id}"
                )
                
                st.session_state.figees[q_id] = True
                st.session_state.reponses_correctes[q_id] = options.index(bonne)

# =============================================
# SECTION 4: FONCTIONS DE GÉNÉRATION
# =============================================
def generer_document(row, template_path):
    """Génération avec gestion des résultats par module"""
    try:
        doc = Document(template_path)
        
        # Initialisation des compteurs par module
        compteurs_modules = {f'mod{i}': 0 for i in range(1, 6)}
        
        # Remplacement des variables simples
        replacements = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }

        # Première passe : remplacement des variables simples dans les paragraphes
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    for run in para.runs:
                        run.text = run.text.replace(key, value)

        # Deuxième passe : traitement des questions
        for q in st.session_state.questions:
            reponses = q['reponses'].copy()
            is_figee = st.session_state.figees.get(q['index'], False)
            
            if is_figee:
                bonne_idx = st.session_state.reponses_correctes.get(q['index'], q.get('correct_idx', 0))
                reponse_correcte = reponses.pop(bonne_idx)
                reponses.insert(0, reponse_correcte)
            else:
                random.shuffle(reponses)
                correct_idx = next((i for i, r in enumerate(reponses) if r['correct']), None)
                if correct_idx is not None:
                    reponse_correcte = reponses.pop(correct_idx)
                    reponses.insert(0, reponse_correcte)

            # Mise à jour du document avec checkbox
            for i, rep in enumerate(reponses):
                para = doc.paragraphs[rep['index']]
                for run in para.runs:
                    checkbox = "☑" if i == 0 else "☐"
                    run.text = f"{rep['lettre']} - {rep['texte']} {checkbox}"

            # Incrémentation du compteur du module
            module = q.get('module', '1')
            if module.isdigit() and 1 <= int(module) <= 5:
                compteurs_modules[f'mod{module}'] += 1

        # Calcul du total général
        total = sum(compteurs_modules.values())
        
        # Création des variables de résultat avec espaces exactes
        replacements.update({
            f'{{{{result_mod{i}}}}}': str(compteurs_modules[f'mod{i}']) 
            for i in range(1, 6)
        })
        replacements['{{result_mod_total}}'] = str(total)

        # Troisième passe : remplacement des variables de résultats dans les paragraphes
        for para in doc.paragraphs:
            for key, value in replacements.items():
                if key in para.text:
                    for run in para.runs:
                        if key in run.text:
                            run.text = run.text.replace(key, str(value))

        # Quatrième passe : remplacement des variables dans les tableaux
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for key, value in replacements.items():
                            if key in para.text:
                                for run in para.runs:
                                    if key in run.text:
                                        # Spécifique pour notre cas : remplacer " " par " " pour correspondre au modèle
                                        run.text = run.text.replace(key, str(value)).replace(' ', ' ')

        return doc
    except Exception as e:
        st.error(f"Erreur de génération : {str(e)}")
        raise

# =============================================
# SECTION 5: GÉNÉRATION PRINCIPALE
# =============================================
if excel_file and word_file and st.session_state.get('questions'):
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
                   
                    for idx, row in df.iterrows():
                        try:
                            doc = generer_document(row, template_path)
                            safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Prénom']))
                            safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(row['Nom']))
                            filename = f"QCM_{safe_prenom}_{safe_nom}.docx"
                            doc.save(os.path.join(tmpdir, filename))
                            zipf.write(os.path.join(tmpdir, filename), filename)
                            progress_bar.progress((idx + 1) / len(df))
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
