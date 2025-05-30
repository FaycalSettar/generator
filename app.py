
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

# — Fonctions utilitaires —

def remplacer_placeholders(paragraph, replacements):
    if not paragraph.text:
        return
    for key, value in replacements.items():
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)
            # gestion des espaces insécables
            ni = key.replace(" ", "\u00a0")
            if ni in run.text:
                run.text = run.text.replace(ni, value)
            # gestion sans espaces
            ns = key.replace(" ", "")
            if ns in run.text:
                run.text = run.text.replace(ns, value)

def detecter_questions(doc):
    questions = []
    current_question = None

    # accepte 1, 1.1, 1.1.2..., doit terminer par ?
    question_pattern = re.compile(r'^\s*(\d+(?:\.\d+)*)\s*[-\s–—.]*\s*(.+?)\s*\?$')
    # réponses de A à D
    answer_pattern   = re.compile(r'^([A-D])\s*[-\s–—.]+\s*(.*?)(\{\{checkbox\}\})?$', re.IGNORECASE)

    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip() \
            .replace("\u00a0", " ") \
            .replace("–", "-") \
            .replace("—", "-")
        if not texte:
            continue

        # DEBUG : voir chaque paragraphe
        st.write(f"PARA {i}: «{texte}»")

        m_q = question_pattern.match(texte)
        if m_q:
            num = m_q.group(1)
            txt = m_q.group(2)
            current_question = {
                "index": i,
                "texte": f"{num} - {txt}?",
                "reponses": [],
                "correct_idx": None,
                "original_text": texte
            }
            questions.append(current_question)
            continue

        if current_question:
            m_r = answer_pattern.match(texte)
            if m_r:
                lettre  = m_r.group(1).upper()
                rep_txt = m_r.group(2).strip()
                is_corr = bool(m_r.group(3))
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": rep_txt,
                    "correct": is_corr,
                    "original_text": texte
                })
                if is_corr:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1

    # filtrer et avertir
    valid = []
    for q in questions:
        if q.get("correct_idx") is not None and len(q["reponses"]) >= 2:
            valid.append(q)
        else:
            st.warning(
                f"Ignorée : {q['texte']} "
                "(bonne réponse manquante ou <2 réponses)"
            )
    return valid

def parse_correct_answers(f):
    if f is None:
        return {}
    try:
        df = pd.read_excel(f)
        df = df.dropna(subset=['Numéro de la question','Réponse correcte'])
        df['Numéro de la question'] = df['Numéro de la question'].astype(str).str.strip()
        df['Réponse correcte']       = df['Réponse correcte'].astype(str).str.strip().str.upper()
        return dict(zip(df['Numéro de la question'], df['Réponse correcte']))
    except Exception as e:
        st.error(f"Erreur lecture corrections : {e}")
        return {}

def calculer_resultat_final(score, total_q=9):
    pct = (score / total_q) * 100
    if pct >= 75:
        return "Acquis"
    elif pct >= 50:
        return "En cours d'acquisition"
    else:
        return "Non acquis"

def generer_document(row, template_bytes):
    try:
        doc = Document(io.BytesIO(template_bytes))
        # placeholders apprenant
        repl = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        # appliquer remplacements
        for p in doc.paragraphs:
            remplacer_placeholders(p, repl)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl)

        # traiter QCM
        corr = st.session_state.get('correct_answers', {})
        score = 0
        for q in st.session_state.questions:
            reps = q['reponses'].copy()
            q_num = q['texte'].split()[0]
            # figé ou non
            if st.session_state.figees.get(q['index'], False):
                bi = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
                cr = reps.pop(bi); reps.insert(0, cr)
            else:
                if q['correct_idx'] is not None:
                    cr = reps.pop(q['correct_idx']); reps.insert(0, cr)
                random.shuffle(reps)
            # écrire réponses dans doc
            for r in reps:
                idx = r['index']
                if idx < len(doc.paragraphs):
                    box = "☑" if reps.index(r) == 0 else "☐"
                    doc.paragraphs[idx].clear()
                    doc.paragraphs[idx].add_run(f"{r['lettre']} - {r['texte']} {box}")
            # scoring
            if q_num in corr and reps[0]['lettre'].upper() == corr[q_num]:
                score += 1

        # résultat final
        res = calculer_resultat_final(score)
        sr = {
            '{{result_mod1}}': str(score),
            '{{result_mod_total}}': str(score),
            '{{result_evaluation}}': res
        }
        for p in doc.paragraphs:
            remplacer_placeholders(p, sr)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, sr)

        return doc, score, res
    except Exception as e:
        st.error(f"Erreur génération doc : {e}")
        st.error(traceback.format_exc())
        return None, 0, "Erreur"

# — Interface Streamlit —

with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    excel_file = st.file_uploader(
        "Excel (Prénom, Nom, Email, Réf Session, Date Évaluation)",
        type="xlsx"
    )
    word_file  = st.file_uploader("Modèle Word .docx", type="docx")
    corr_file  = st.file_uploader("Réponses correctes (xlsx)", type="xlsx")

# initialiser session
for key in ('questions','figees','reponses_correctes'):
    if key not in st.session_state:
        st.session_state[key] = [] if key=='questions' else {}
if 'current_template' not in st.session_state:
    st.session_state.current_template = None
if 'doc_template' not in st.session_state:
    st.session_state.doc_template = None

# charger Word & détecter questions
if word_file and st.session_state.current_template != word_file.name:
    try:
        data = word_file.getvalue()
        doc = Document(io.BytesIO(data))
        st.session_state.doc_template = data
        qs = detecter_questions(doc)
        st.session_state.questions = qs
        st.session_state.current_template = word_file.name
        if qs:
            st.success(f"✅ {len(qs)} questions détectées")
            with st.expander("🔍 Questions détectées", expanded=True):
                for idx, q in enumerate(qs, 1):
                    st.write(f"**{idx}. {q['texte']}**")
                    for j, r in enumerate(q['reponses']):
                        mark = "✅" if j == q['correct_idx'] else "☐"
                        st.write(f"{mark} {r['lettre']} - {r['texte']}")
        else:
            st.warning("⚠️ Aucune question détectée. Vérifiez le format.")
    except Exception as e:
        st.error(f"Erreur chargement Word : {e}")
        st.error(traceback.format_exc())

# charger corrections
if corr_file:
    st.session_state.correct_answers = parse_correct_answers(corr_file)
    st.success(f"✅ {len(st.session_state.correct_answers)} corrections chargées")

# configuration questions figées
if st.session_state.questions:
    st.markdown("### Configuration des questions")
    for q in st.session_state.questions:
        with st.expander(q['texte'], expanded=False):
            fig = st.checkbox("Figer cette question", key=f"figer_{q['index']}")
            if fig:
                opts = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                default = q['correct_idx'] or 0
                choix   = st.selectbox("Bonne réponse", opts, index=default, key=f"bonne_{q['index']}")
                st.session_state.figees[q['index']] = True
                st.session_state.reponses_correctes[q['index']] = opts.index(choix)

# génération QCM
if excel_file and st.session_state.questions and st.button("Générer les QCM"):
    try:
        df = pd.read_excel(excel_file)
        need = ['Prénom','Nom','Email','Référence Session','Date Évaluation']
        miss = [c for c in need if c not in df.columns]
        if miss:
            st.error(f"Colonnes manquantes : {miss}")
            st.stop()

        buf  = io.BytesIO()
        recap = []
        with ZipFile(buf, 'w') as zf:
            prog  = st.progress(0)
            total = len(df)
            for i, row in df.iterrows():
                doc_out, sc, re = generer_document(row, st.session_state.doc_template)
                if doc_out:
                    recap.append({
                        "Prénom": row["Prénom"],
                        "Nom": row["Nom"],
                        "Réf": row["Référence Session"],
                        "Score": sc,
                        "Résultat": re
                    })
                    fn = f"QCM_{row['Prénom']}_{row['Nom']}.docx"
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
                    doc_out.save(tmp.name)
                    zf.write(tmp.name, fn)
                prog.progress((i+1)/total)

            if recap:
                df_r = pd.DataFrame(recap)
                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                df_r.to_excel(tmp2.name, index=False)
                zf.write(tmp2.name, "Recapitulatif_QCM.xlsx")

        buf.seek(0)
        st.success("✅ Génération terminée")
        st.download_button(
            "⬇️ Télécharger ZIP", data=buf,
            file_name="QCM_Personnalises.zip",
            mime="application/zip"
        )
    except Exception as e:
        st.error(f"ERREUR critique : {e}")
        st.error(traceback.format_exc())

# légende résultats
st.markdown("### Légende résultats")
st.info("""
- **Acquis** : ≥ 75%  
- **En cours d'acquisition** : 50–75%  
- **Non acquis** : < 50%
""")
