import streamlit as st
import pandas as pd
from docx import Document
import random
import os
import tempfile
from zipfile import ZipFile
import re
import traceback
from collections import defaultdict
import unicodedata

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# =============================================
# 1) FONCTIONS UTILITAIRES
# =============================================
def normalize_text(text):
    """Normalise le texte pour les comparaisons"""
    return unicodedata.normalize('NFKD', str(text)).encode('ASCII', 'ignore').decode().lower().strip()

def detecter_questions(doc):
    """Détection améliorée des questions dans le document Word"""
    questions = []
    current = None
    quest_pat = re.compile(
        r'^(\d+\.\d+)'          # Numéro de question
        r'[\s\-–—)]*'            # Séparateurs
        r'\s*(.+?)'              # Texte question
        r'\s*\??'                # Point d'interrogation optionnel
        r'$', 
        flags=re.IGNORECASE
    )
    
    rep_pat = re.compile(
        r'^([A-D])'              # Lettre réponse
        r'[\s\-–—).]*'           # Séparateurs
        r'\s*(.+?)'              # Texte réponse
        r'(\s*\{\{checkbox\}\})?$'  # Checkbox
    )

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        
        # Détection question
        m_quest = quest_pat.match(text)
        if m_quest:
            current = {
                'qnum': m_quest.group(1),
                'index': i,
                'texte': f"{m_quest.group(1)} - {m_quest.group(2)}",
                'reponses': [],
                'correct_idx': None,
                'module': m_quest.group(1).split('.')[0]
            }
            questions.append(current)
            continue
        
        # Détection réponse
        if current:
            m_rep = rep_pat.match(text)
            if m_rep:
                reponse = {
                    'index': i,
                    'lettre': m_rep.group(1).upper(),
                    'texte': m_rep.group(2).strip(),
                    'correct': False
                }
                current['reponses'].append(reponse)
                
    return questions

# =============================================
# 2) UPLOAD DES FICHIERS
# =============================================
with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    excel_file = st.file_uploader(
        "Fichier Excel (Prénom, Nom, Email, Référence Session, Date Évaluation)",
        type="xlsx"
    )
    word_file = st.file_uploader("Modèle Word (.docx)", type="docx")
    correction_file = st.file_uploader("Fichier Quizz.xlsx (corrections + résultats)", type="xlsx")

# =============================================
# 3) TRAITEMENT DE LA CORRECTION
# =============================================
def process_correction(corr_df, questions):
    """Traite le fichier de correction et calcule les résultats"""
    results = defaultdict(int)
    
    # Vérification format automatique
    if {'module', 'bonnes_reponses'}.issubset(corr_df.columns.str.lower()):
        for _, row in corr_df.iterrows():
            mod = str(row['module']).strip()
            results[mod] += int(row['bonnes_reponses'])
        total = sum(results.values())
        return dict(results), total
    
    # Mode manuel
    corr_map = {}
    for _, row in corr_df.iterrows():
        qnum = str(row['Numéro de la question']).strip()
        rep = str(row['Réponse correcte']).strip().upper()
        if qnum and rep:
            corr_map[qnum] = rep

    # Validation complétude
    missing = []
    for q in questions:
        qnum = q['qnum']
        if qnum not in corr_map:
            missing.append(qnum)
            continue
            
        rep_correct = corr_map[qnum]
        for idx, r in enumerate(q['reponses']):
            if r['lettre'] == rep_correct:
                results[q['module']] += 1
                break
        else:
            st.error(f"Réponse invalide {rep_correct} pour question {qnum}")

    if missing:
        st.error(f"Questions sans correction: {', '.join(missing)}")
        st.stop()

    total = sum(results.values())
    return dict(results), total

# =============================================
# 4) INTERFACE UTILISATEUR
# =============================================
if all([excel_file, word_file, correction_file]):
    if 'results_mod' not in st.session_state:
        try:
            # Chargement des données
            df = pd.read_excel(excel_file)
            corr_df = pd.read_excel(correction_file)
            doc = Document(word_file)
            
            # Vérification colonnes
            required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
            df_cols_norm = [normalize_text(c) for c in df.columns]
            missing = [c for c in required_cols if normalize_text(c) not in df_cols_norm]
            
            if missing:
                st.error(f"Colonnes manquantes: {', '.join(missing)}")
                st.stop()

            # Détection questions
            questions = detecter_questions(doc)
            if not questions:
                st.error("Aucune question détectée dans le modèle Word!")
                st.stop()

            # Calcul résultats
            results_mod, total = process_correction(corr_df, questions)
            
            # Initialisation session
            st.session_state.update({
                'questions': questions,
                'results_mod': results_mod,
                'results_total': total,
                'figees': {},
                'reponses_correctes': {}
            })
            
        except Exception as e:
            st.error(f"Erreur de traitement: {str(e)}")
            st.text(traceback.format_exc())
            st.stop()

# =============================================
# 5) GESTION DES QUESTIONS FIGÉES
# =============================================
if 'questions' in st.session_state:
    st.markdown("### Étape 2 : Configuration des questions")
    for q in st.session_state.questions:
        col1, col2 = st.columns([1, 5])
        with col1:
            froze = st.checkbox(
                f"Q{q['qnum']}", 
                key=f"fig_{q['index']}",
                help=q['texte']
            )
            st.session_state.figees[q['index']] = froze
        with col2:
            if froze:
                options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                default = next((i for i, r in enumerate(q['reponses']) if r['correct']), 0)
                new_correct = st.selectbox(
                    f"Réponse correcte Q{q['qnum']}",
                    options,
                    index=default,
                    key=f"rep_{q['index']}"
                )
                st.session_state.reponses_correctes[q['index']] = options.index(new_correct)

# =============================================
# 6) GÉNÉRATION DES DOCUMENTS
# =============================================
def generer_document(row, tpl_path):
    """Génère un document personnalisé"""
    doc = Document(tpl_path)
    
    # Mapping des données
    replacements = {
        '{{prenom}}': row['Prénom'],
        '{{nom}}': row['Nom'],
        '{{email}}': row['Email'],
        '{{ref_session}}': row['Référence Session'],
        '{{date_evaluation}}': row['Date Évaluation']
    }
    
    # Remplacement des variables
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if k in p.text:
                for run in p.runs:
                    run.text = run.text.replace(k, str(v))
    
    # Gestion des réponses
    for q in st.session_state.questions:
        reps = q['reponses'].copy()
        
        # Réorganisation des réponses
        if st.session_state.figees.get(q['index'], False):
            correct_idx = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
            bonne_reponse = reps.pop(correct_idx)
            reps.insert(0, bonne_reponse)
        else:
            random.shuffle(reps)
        
        # Écriture des réponses
        for r in reps:
            try:
                p = doc.paragraphs[r['index']]
                case = "☑" if r['correct'] else "☐"
                new_text = f"{r['lettre']} - {r['texte']}   {case}"
                
                if p.runs:
                    p.runs[0].text = new_text
                    for run in p.runs[1:]:
                        run.text = ''
                else:
                    p.text = new_text
            except IndexError:
                continue

    # Résultats par module
    for mod, score in st.session_state.results_mod.items():
        replace_in_doc(doc, f'{{{{result_mod{mod}}}}}', str(score))
    
    # Score total
    replace_in_doc(doc, '{{result_mod_total}}', str(st.session_state.results_total))
    
    return doc

# =============================================
# 7) GÉNÉRATION FINALE
# =============================================
if all([excel_file, word_file, correction_file]):
    st.markdown("---")
    if st.button("Générer tous les QCM"):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Préparation template
                tpl_path = os.path.join(tmpdir, 'template.docx')
                with open(tpl_path, 'wb') as f:
                    f.write(word_file.getbuffer())
                
                # Création archive
                zip_path = os.path.join(tmpdir, 'QCM_Resultats.zip')
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    df = pd.read_excel(excel_file)
                    
                    for i, row in df.iterrows():
                        try:
                            doc = generer_document(row, tpl_path)
                            filename = f"QCM_{row['Prénom']}_{row['Nom']}.docx".replace(' ', '_')
                            doc_path = os.path.join(tmpdir, filename)
                            doc.save(doc_path)
                            zipf.write(doc_path, arcname=filename)
                        except Exception as e:
                            st.error(f"Erreur avec {row['Prénom']} {row['Nom']}: {str(e)}")
                        progress_bar.progress((i + 1) / len(df))
                
                # Téléchargement
                with open(zip_path, 'rb') as f:
                    st.success("Génération terminée avec succès!")
                    st.download_button(
                        "📥 Télécharger l'archive",
                        data=f.read(),
                        file_name="QCM_Resultats.zip",
                        mime="application/zip"
                    )
        except Exception as e:
            st.error(f"Erreur critique: {str(e)}")
            st.text(traceback.format_exc())
