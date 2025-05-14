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
# SECTION 1 : UPLOAD DES FICHIERS
# =============================================
with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    excel_file      = st.file_uploader("Fichier Excel (%) colonnes : Prénom, Nom, Email, Référence Session, Date Évaluation", type="xlsx")
    word_file       = st.file_uploader("Modèle Word (.docx)", type="docx")
    correction_file = st.file_uploader("Fichier de correction (Quizz.xlsx)", type="xlsx")

# =============================================
# SECTION 2 : DÉTECTION DES QUESTIONS & CORRECTION
# =============================================
def detecter_questions(doc: Document):
    """Repère toutes les questions (groupes 1.1, 1.2, …) et leurs réponses brutes."""
    questions = []
    current = None
    # regex question et réponse
    quest_pat = re.compile(r'^(\d+\.\d+)\s*[-–—)]?\s*(.+?)\?$')
    rep_pat   = re.compile(r'^([A-D])[\s\-–—).]+\s*(.+)$')
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
                lettre = mr.group(1)
                txt    = mr.group(2).strip()
                current["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": txt,
                    # correct sera mis à jour juste après
                    "correct": False
                })
    return questions

if word_file and correction_file and excel_file:
    # Initialisation une seule fois
    if "questions" not in st.session_state:
        # 1) Détection brute
        doc0 = Document(word_file)
        raw_q = detecter_questions(doc0)

        # 2) Chargement des bonnes réponses
        corr_df = pd.read_excel(correction_file)
        # On suppose qu'il y a une colonne "Numéro de la question" et "Réponse correcte"
        corr_map = {
            str(r["Numéro de la question"]): r["Réponse correcte"].strip().upper()
            for _, r in corr_df.iterrows()
        }

        # 3) Appliquer la correction sur raw_q
        questions = []
        for q in raw_q:
            correct_letter = corr_map.get(q["qnum"])
            for idx, rep in enumerate(q["reponses"]):
                rep["correct"] = (rep["lettre"].upper() == correct_letter)
                if rep["correct"]:
                    q["correct_idx"] = idx
            # On garde seulement si on a trouvé une bonne réponse
            if q["correct_idx"] is not None:
                questions.append(q)
        st.session_state.questions = questions

        # 4) Calcul des résultats par module
        results_mod = {}
        for q in questions:
            mod = q["qnum"].split(".")[0]
            results_mod[mod] = results_mod.get(mod, 0) + 1
        total = sum(results_mod.values())
        st.session_state.results_mod   = results_mod
        st.session_state.results_total = total

        # 5) Initialisation du figement (optionnel)
        st.session_state.figees             = {}
        st.session_state.reponses_correctes = {}

# =============================================
# SECTION 3 : CONFIGURATION DES QUESTIONS (FIGER)
# =============================================
if "questions" in st.session_state:
    st.markdown("### Étape 2 : configuration des QCM (figer les questions si besoin)")
    for q in st.session_state.questions:
        qid = q["index"]
        num = q["qnum"]
        c1, c2 = st.columns([1,5])
        with c1:
            fig = st.checkbox(f"Q{num}", key=f"fig_{qid}", help=q["texte"])
            st.session_state.figees[qid] = fig
        with c2:
            if fig:
                opts = [f"{r['lettre']} - {r['texte']}" for r in q["reponses"]]
                default = q["correct_idx"]
                sel = st.selectbox(f"Bonne réponse Q{num}", opts, index=default, key=f"sel_{qid}")
                st.session_state.reponses_correctes[qid] = opts.index(sel)

# =============================================
# SECTION 4 : GÉNÉRATION D’UN DOC POUR UNE LIGNE
# =============================================
def generer_document(row, tpl_path):
    doc = Document(tpl_path)
    # → Remplacer les variables de l’entête
    mapping = {
        "{{prenom}}":        str(row["Prénom"]),
        "{{nom}}":           str(row["Nom"]),
        "{{email}}":         str(row["Email"]),
        "{{ref_session}}":   str(row["Référence Session"]),
        "{{date_evaluation}}": str(row["Date Évaluation"])
    }
    for p in doc.paragraphs:
        for k, v in mapping.items():
            p.text = p.text.replace(k, v)

    # → Traiter chaque question
    for q in st.session_state.questions:
        reps = q["reponses"].copy()
        idx_q = q["index"]
        # figée ?
        if st.session_state.figees.get(idx_q, False):
            ci = st.session_state.reponses_correctes.get(idx_q, q["correct_idx"])
            bonne = reps.pop(ci)
            reps.insert(0, bonne)
        else:
            random.shuffle(reps)
            # s'assurer que la bonne reste incluse
            # (elle y est, car rep.correct définit q["correct_idx"])

        # écrire cases
        for pos, r in enumerate(reps):
            p = doc.paragraphs[r["index"]]
            case = "☑" if r["correct"] else "☐"
            p.text = f"{r['lettre']} - {r['texte']}   {case}"

    # → Remplacer les placeholders de résultats
    # Modules 1→5
    for m, count in st.session_state.results_mod.items():
        token = f"{{{{result_mod{m}}}}}"
        for p in doc.paragraphs:
            p.text = p.text.replace(token, str(count))
    # Total
    for p in doc.paragraphs:
        p.text = p.text.replace("{{result_mod_total}}", str(st.session_state.results_total))

    return doc

# =============================================
# SECTION 5 : GÉNÉRATION GLOBALE & ZIP
# =============================================
if excel_file and word_file and correction_file and st.session_state.get("questions"):
    st.markdown("---")
    if st.button("Générer les QCM et résultats"):
        try:
            df = pd.read_excel(excel_file)
            # Vérifier colonnes
            need = ["Prénom","Nom","Email","Référence Session","Date Évaluation"]
            miss = [c for c in need if c not in df.columns]
            if miss:
                st.error("Colonnes manquantes : " + ", ".join(miss))
                st.stop()

            with tempfile.TemporaryDirectory() as td:
                # sauvegarde template
                tpl = os.path.join(td, "template.docx")
                with open(tpl, "wb") as f:
                    f.write(word_file.getbuffer())

                # créer ZIP
                zip_path = os.path.join(td, "QCM_Personnalises.zip")
                with ZipFile(zip_path, "w") as z:
                    pb = st.progress(0)
                    for i, row in df.iterrows():
                        try:
                            doc = generer_document(row, tpl)
                            # nom de fichier safe
                            pren = re.sub(r"[^A-Za-z0-9]","_", str(row["Prénom"]))
                            nomf = re.sub(r"[^A-Za-z0-9]","_", str(row["Nom"]))
                            fn = f"QCM_{pren}_{nomf}.docx"
                            outp = os.path.join(td, fn)
                            doc.save(outp)
                            z.write(outp, fn)
                        except Exception as e:
                            st.error(f"Échec {row['Prénom']} {row['Nom']} : {e}")
                        pb.progress((i+1)/len(df))

                with open(zip_path, "rb") as f:
                    st.success("✅ Génération des QCM + résultats terminée !")
                    st.download_button(
                        "Télécharger l'archive ZIP",
                        f, file_name="QCM_Résultats.zip", mime="application/zip"
                    )

        except Exception as e:
            st.error("ERREUR FATALE : " + str(e))
            st.text(traceback.format_exc())
