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

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# =============================================
# 1) UPLOAD
# =============================================
with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    excel_file = st.file_uploader(
        "Fichier Excel (Prénom, Nom, Email, Référence Session, Date Évaluation)",
        type="xlsx"
    )
    word_file = st.file_uploader("Modèle Word (.docx)", type="docx")
    correction_file = st.file_uploader("Fichier Quizz.xlsx (corrections + résultats)", type="xlsx")

# =============================================
# 2) DÉTECTION DES QUESTIONS DANS LE WORD
# =============================================
def detecter_questions_ameliore(doc: Document):
    questions = []
    current = None
    quest_pat = re.compile(
        r'^(\d+\.\d+)'          # Numéro de question
        r'[\s\-–—)]*'           # Séparateurs optionnels
        r'\s*(.+?\??)'          # Texte de la question
        r'$', 
        flags=re.IGNORECASE
    )
    
    rep_pat = re.compile(
        r'^([A-D])'             # Lettre réponse
        r'[\s\-–—).]*'          # Séparateurs
        r'\s*(.+?)'             # Texte réponse
        r'(\s*\{\{checkbox\}\})?$'  # Marqueur checkbox
    )

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        
        # Détection question
        m_quest = quest_pat.match(text)
        if m_quest:
            current = {
                "qnum": m_quest.group(1),
                "index": i,
                "texte": f"{m_quest.group(1)} - {m_quest.group(2)}",
                "reponses": [],
                "correct_idx": None,
                "module": m_quest.group(1).split('.')[0]
            }
            questions.append(current)
            continue
        
        # Détection réponse
        if current:
            m_rep = rep_pat.match(text)
            if m_rep:
                reponse = {
                    "index": i,
                    "lettre": m_rep.group(1).upper(),
                    "texte": m_rep.group(2).strip(),
                    "correct": False
                }
                current["reponses"].append(reponse)
                
    return questions

# =============================================
# 3) TRAITEMENT DE LA CORRECTION
# =============================================
def process_correction(corr_df, questions):
    # Vérification format correction
    if {"Module", "Nombre de bonnes réponses"}.issubset(corr_df.columns):
        results_mod = corr_df.groupby("Module")["Nombre de bonnes réponses"].sum().astype(int).to_dict()
        total = sum(results_mod.values())
        return results_mod, total
    
    # Mode manuel si colonnes manquantes
    corr_map = {}
    for _, row in corr_df.iterrows():
        qnum = str(row["Numéro de la question"]).strip()
        rep = str(row["Réponse correcte"]).strip().upper()
        if qnum and rep:
            corr_map[qnum] = rep

    # Validation complétude des corrections
    missing = []
    results_mod = defaultdict(int)
    for q in questions:
        if q["qnum"] not in corr_map:
            missing.append(q["qnum"])
        else:
            if corr_map[q["qnum"]] in [r["lettre"] for r in q["reponses"]]:
                results_mod[q["module"]] += 1
            else:
                st.error(f"Réponse invalide {corr_map[q['qnum']]} pour question {q['qnum']}")

    if missing:
        st.error(f"Questions sans correction: {', '.join(missing)}")
        st.stop()

    total = sum(results_mod.values())
    return dict(results_mod), total

# =============================================
# 4) INTERFACE UTILISATEUR
# =============================================
if excel_file and word_file and correction_file:
    if "results_mod" not in st.session_state:
        try:
            # Chargement fichiers
            corr_df = pd.read_excel(correction_file)
            doc = Document(word_file)
            questions = detecter_questions_ameliore(doc)
            
            # Validation questions détectées
            if not questions:
                st.error("Aucune question détectée dans le modèle Word!")
                st.stop()

            # Traitement correction
            results_mod, total = process_correction(corr_df, questions)
            
            # Enregistrement état
            st.session_state.update({
                "questions": questions,
                "results_mod": results_mod,
                "results_total": total,
                "figees": {},
                "reponses_correctes": {}
            })
            
        except Exception as e:
            st.error(f"Erreur de traitement: {str(e)}")
            st.text(traceback.format_exc())
            st.stop()

# =============================================
# 5) GESTION DES QUESTIONS FIGÉES
# =============================================
if "questions" in st.session_state:
    st.markdown("### Étape 2 : Configuration des questions")
    for q in st.session_state.questions:
        col1, col2 = st.columns([1, 5])
        with col1:
            froze = st.checkbox(
                f"Q{q['qnum']}", 
                key=f"fig_{q['index']}",
                help=q["texte"]
            )
            st.session_state.figees[q["index"]] = froze
        with col2:
            if froze:
                options = [f"{r['lettre']} - {r['texte']}" for r in q["reponses"]]
                default = next((i for i, r in enumerate(q["reponses"]) if r["correct"]), 0)
                new_correct = st.selectbox(
                    f"Réponse correcte Q{q['qnum']}",
                    options,
                    index=default,
                    key=f"rep_{q['index']}"
                )
                st.session_state.reponses_correctes[q["index"]] = options.index(new_correct)

# =============================================
# 6) GÉNÉRATION DES DOCUMENTS
# =============================================
def generer_document(row, tpl_path):
    doc = Document(tpl_path)
    
    # Remplacement des variables utilisateur
    replacements = {
        "{{prenom}}": row["Prénom"],
        "{{nom}}": row["Nom"],
        "{{email}}": row["Email"],
        "{{ref_session}}": row["RÃ©fÃ©rence Session"],  # Attention à l'encodage
        "{{date_evaluation}}": row["Date Ã‰valuation"]
    }
    
    # Remplacement dans tout le document
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if k in p.text:
                for run in p.runs:
                    run.text = run.text.replace(k, str(v))
    
    # Remplacement dans les tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in replacements.items():
                        if k in p.text:
                            for run in p.runs:
                                run.text = run.text.replace(k, str(v))
                                
    # Gestion des questions et réponses
    for q in st.session_state.questions:
        reps = q["reponses"].copy()
        
        if st.session_state.figees.get(q["index"], False):
            correct_idx = st.session_state.reponses_correctes.get(q["index"], q["correct_idx"])
            bonne_reponse = reps.pop(correct_idx)
            reps.insert(0, bonne_reponse)
        else:
            random.shuffle(reps)
        
        for r in reps:
            try:
                p = doc.paragraphs[r["index"]]
                case = "☑" if r["correct"] else "☐"
                new_text = f"{r['lettre']} - {r['texte']}   {case}"
                
                # Conservation du style original
                if p.runs:
                    p.runs[0].text = new_text
                    for run in p.runs[1:]:
                        run.text = ""
                else:
                    p.text = new_text
            except IndexError:
                st.error(f"Erreur d'index pour la question {q['qnum']}")

    # Remplissage des résultats
    for mod in st.session_state.results_mod:
        replace_in_doc(doc, f"{{{{result_mod{mod}}}}}", str(st.session_state.results_mod[mod]))
    replace_in_doc(doc, "{{result_mod_total}}", str(st.session_state.results_total))
    
    return doc

# =============================================
# 7) GÉNÉRATION FINALE
# =============================================
if excel_file and word_file and correction_file:
    st.markdown("---")
    if st.button("Générer tous les QCM"):
        try:
            df = pd.read_excel(excel_file)
            required_columns = ["Prénom", "Nom", "Email", "RÃ©fÃ©rence Session", "Date Ã‰valuation"]
            if not all(col in df.columns for col in required_columns):
                missing = [col for col in required_columns if col not in df.columns]
                st.error(f"Colonnes manquantes: {', '.join(missing)}")
                st.stop()

            with tempfile.TemporaryDirectory() as tmpdir:
                # Sauvegarde du template
                tpl_path = os.path.join(tmpdir, "template.docx")
                with open(tpl_path, "wb") as f:
                    f.write(word_file.getbuffer())
                
                # Création ZIP
                zip_path = os.path.join(tmpdir, "QCM_Resultats.zip")
                with ZipFile(zip_path, "w") as zipf:
                    progress_bar = st.progress(0)
                    total_users = len(df)
                    
                    for i, row in df.iterrows():
                        try:
                            doc = generer_document(row, tpl_path)
                            filename = f"QCM_{row['Prénom']}_{row['Nom']}.docx".replace(" ", "_")
                            doc_path = os.path.join(tmpdir, filename)
                            doc.save(doc_path)
                            zipf.write(doc_path, arcname=filename)
                        except Exception as e:
                            st.error(f"Erreur avec {row['Prénom']} {row['Nom']}: {str(e)}")
                        progress_bar.progress((i + 1) / total_users)
                
                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.success("Génération terminée avec succès!")
                    st.download_button(
                        "Télécharger les QCM",
                        data=f.read(),
                        file_name="QCM_Resultats.zip",
                        mime="application/zip"
                    )
        except Exception as e:
            st.error(f"Erreur critique: {str(e)}")
            st.text(traceback.format_exc())
