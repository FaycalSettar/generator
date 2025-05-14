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
# 1) UPLOAD
# =============================================
with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    excel_file      = st.file_uploader(
        "Fichier Excel (Prénom, Nom, Email, Référence Session, Date Évaluation)",
        type="xlsx"
    )
    word_file       = st.file_uploader("Modèle Word (.docx)", type="docx")
    correction_file = st.file_uploader("Fichier de correction (Quizz.xlsx)", type="xlsx")

# =============================================
# 2) DÉTECTION DES QUESTIONS & CORRECTION
# =============================================
def detecter_questions(doc: Document):
    """Repère questions numérotées et collecte leurs réponses."""
    questions = []
    current = None
    quest_pat = re.compile(r'^(\d+\.\d+)\s*[-–—)]?\s*(.+?)\?$')
    rep_pat   = re.compile(r'^([A-D])[\s\-–—).]+\s*(.+?)\s*({{checkbox}})?$')
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        m = quest_pat.match(t)
        if m:
            current = {
                "qnum":       m.group(1),
                "index":      i,
                "texte":      f"{m.group(1)} - {m.group(2)}?",
                "reponses":   [],
                "correct_idx": None
            }
            questions.append(current)
            continue
        if current:
            mr = rep_pat.match(t)
            if mr:
                lettre     = mr.group(1)
                texte_rep  = mr.group(2).strip()
                current["reponses"].append({
                    "index":  i,
                    "lettre": lettre,
                    "texte":  texte_rep,
                    "correct": False
                })
    return questions

if excel_file and word_file and correction_file:
    if "questions" not in st.session_state:
        # 1. détecter questions
        doc0   = Document(word_file)
        raw_q  = detecter_questions(doc0)
        # 2. charger correction depuis Quizz.xlsx
        corr_df = pd.read_excel(correction_file)
        corr_map = {
            str(r["Numéro de la question"]): r["Réponse correcte"].strip().upper()
            for _, r in corr_df.iterrows()
        }
        # 3. appliquer correction
        questions = []
        for q in raw_q:
            lettre_ok = corr_map.get(q["qnum"])
            for idx, rep in enumerate(q["reponses"]):
                if rep["lettre"] == lettre_ok:
                    rep["correct"] = True
                    q["correct_idx"] = idx
            if q["correct_idx"] is not None:
                questions.append(q)
        st.session_state.questions = questions
        # 4. calcul résultat par module (nombre de bonnes réponses)
        results_mod = {}
        for q in questions:
            mod = q["qnum"].split(".")[0]
            results_mod[mod] = results_mod.get(mod, 0) + 1
        st.session_state.results_mod   = results_mod
        st.session_state.results_total = sum(results_mod.values())
        # 5. préparation du figement
        st.session_state.figees = {}
        st.session_state.reponses_correctes = {}

# =============================================
# 3) CONFIGURATION (FIGER)
# =============================================
if "questions" in st.session_state:
    st.markdown("### Étape 2 : Optionnel – figer certaines questions")
    for q in st.session_state.questions:
        qid  = q["index"]
        num  = q["qnum"]
        c1, c2 = st.columns([1,5])
        with c1:
            f = st.checkbox(f"Q{num}", key=f"fig_{qid}", help=q["texte"])
            st.session_state.figees[qid] = f
        with c2:
            if f:
                opts = [f"{r['lettre']} - {r['texte']}" for r in q["reponses"]]
                default = q["correct_idx"]
                sel     = st.selectbox(f"Bonne réponse Q{num}", opts, index=default, key=f"sel_{qid}")
                st.session_state.reponses_correctes[qid] = opts.index(sel)

# =============================================
# 4) FONCTION DE REMPLACEMENT DANS LE DOC
# =============================================
def replace_in_doc(doc, token, value):
    # Paragraphes
    for p in doc.paragraphs:
        if token in p.text:
            p.text = p.text.replace(token, value)
    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if token in p.text:
                        p.text = p.text.replace(token, value)

# =============================================
# 5) GÉNÉRATION D’UN DOC POUR UNE LIGNE
# =============================================
def generer_document(row, tpl_path):
    doc = Document(tpl_path)
    # → remplacer les variables de l’en-tête
    vars_map = {
        "{{prenom}}":         str(row["Prénom"]),
        "{{nom}}":            str(row["Nom"]),
        "{{email}}":          str(row["Email"]),
        "{{ref_session}}":    str(row["Référence Session"]),
        "{{date_evaluation}}":str(row["Date Évaluation"])
    }
    for p in doc.paragraphs:
        for k, v in vars_map.items():
            if k in p.text:
                p.text = p.text.replace(k, v)
    # → traiter chaque question
    for q in st.session_state.questions:
        reps = q["reponses"].copy()
        idx = q["index"]
        if st.session_state.figees.get(idx, False):
            ci = st.session_state.reponses_correctes.get(idx, q["correct_idx"])
            bonne = reps.pop(ci)
            reps.insert(0, bonne)
        else:
            random.shuffle(reps)
        for pos, r in enumerate(reps):
            p = doc.paragraphs[r["index"]]
            case = "☑" if r["correct"] else "☐"
            p.text = f"{r['lettre']} - {r['texte']}   {case}"
    # → remplacer les placeholders de résultats
    for m, cnt in st.session_state.results_mod.items():
        replace_in_doc(doc, f"{{{{result_mod{m}}}}}", str(cnt))
    replace_in_doc(doc, "{{result_mod_total}}", str(st.session_state.results_total))
    return doc

# =============================================
# 6) BOUTON & ZIP FINAL
# =============================================
if excel_file and word_file and correction_file and st.session_state.get("questions"):
    st.markdown("---")
    if st.button("Générer tous les QCM + résultats"):
        try:
            df = pd.read_excel(excel_file)
            needed = ["Prénom","Nom","Email","Référence Session","Date Évaluation"]
            missing = [c for c in needed if c not in df.columns]
            if missing:
                st.error("Colonnes manquantes : " + ", ".join(missing))
                st.stop()
            with tempfile.TemporaryDirectory() as td:
                tpl = os.path.join(td, "template.docx")
                with open(tpl, "wb") as f:
                    f.write(word_file.getbuffer())
                zip_path = os.path.join(td, "QCM_Resultats.zip")
                with ZipFile(zip_path, "w") as z:
                    pb = st.progress(0)
                    for i, row in df.iterrows():
                        try:
                            doc = generer_document(row, tpl)
                            pren = re.sub(r"[^A-Za-z0-9]","_", str(row["Prénom"]))
                            nomf = re.sub(r"[^A-Za-z0-9]","_", str(row["Nom"]))
                            fn   = f"QCM_{pren}_{nomf}.docx"
                            outp = os.path.join(td, fn)
                            doc.save(outp)
                            z.write(outp, fn)
                        except Exception as e:
                            st.error(f"Échec {row['Prénom']} {row['Nom']} : {e}")
                        pb.progress((i+1)/len(df))
                with open(zip_path, "rb") as f:
                    st.success("✅ Génération terminée avec succès !")
                    st.download_button(
                        "Télécharger l'archive ZIP",
                        data=f,
                        file_name="QCM_Resultats.zip",
                        mime="application/zip"
                    )
        except Exception as e:
            st.error("ERREUR FATALE : " + str(e))
            st.text(traceback.format_exc())
