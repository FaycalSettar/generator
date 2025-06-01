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
    """
    Cette version de detecter_questions reconnaît :
      1) les questions numérotées (1, 1.1, etc.) se terminant par un '?'
      2) les questions non numérotées commençant par un tiret ('-','–','—') et finissant par '?'
    Chaque question valide doit avoir au moins deux réponses A–D et une réponse marquée avec '{{checkbox}}'.
    """
    questions = []
    current_question = None
    compteur_non_numerote = 0

    # 1) Questions numérotées, ex. "1.1 - Texte ?"
    pattern_num = re.compile(r'^\s*(\d+(?:\.\d+)*)\s*[-\s–—.]*\s*(.+?)\s*\?$')
    # 2) Questions non numérotées, ex. "- Texte ?"
    pattern_non_num = re.compile(r'^\s*[-–—]\s*(.+?)\s*\?$')

    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip() \
                    .replace("\u00a0", " ") \
                    .replace("–", "-") \
                    .replace("—", "-")
        if not texte:
            continue

        # 1) Tentative de match pour question numérotée
        m_num = pattern_num.match(texte)
        if m_num:
            num = m_num.group(1)
            txt = m_num.group(2)
            libelle = f"{num} - {txt}?"
            current_question = {
                "index": i,
                "texte": libelle,
                "numero": num,
                "reponses": [],
                "correct_idx": None,
                "original_text": texte
            }
            questions.append(current_question)
            continue

        # 2) Tentative de match pour question non numérotée
        m_non = pattern_non_num.match(texte)
        if m_non:
            compteur_non_numerote += 1
            num_fake = f"NN{compteur_non_numerote}"
            txt = m_non.group(1)
            libelle = f"{num_fake} - {txt}?"
            current_question = {
                "index": i,
                "texte": libelle,
                "numero": num_fake,
                "reponses": [],
                "correct_idx": None,
                "original_text": texte
            }
            questions.append(current_question)
            continue

        # 3) Si on est dans une question en cours, vérifier une réponse A–D
        if current_question:
            m_ans = re.match(r'^([A-D])\s*[-\s–—.]+\s*(.*?)\s*(\{\{checkbox\}\})?$', texte, re.IGNORECASE)
            if m_ans:
                lettre   = m_ans.group(1).upper()
                rep_txt  = m_ans.group(2).strip()
                is_corr  = bool(m_ans.group(3))
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": rep_txt,
                    "correct": is_corr,
                    "original_text": texte
                })
                if is_corr:
                    current_question["correct_idx"] = len(current_question["reponses"]) - 1
                continue

    # Filtrer les questions valides (>=2 réponses et au moins une correcte)
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

def calculer_resultat_final(score, total_q):
    """
    Renvoie le libellé ("Acquis", "En cours d'acquisition", "Non acquis")
    en fonction du pourcentage score/total_q.
    """
    pct = (score / total_q) * 100 if total_q > 0 else 0
    if pct >= 75:
        return "Acquis : ≥ 75%"
    elif pct >= 50:
        return "En cours d'acquisition : 50–75%"
    else:
        return "Non acquis : < 50%"

def generer_document(row, template_bytes):
    try:
        doc = Document(io.BytesIO(template_bytes))
        # placeholders apprenant
        # Gestion de la date au format JJ/MM/AAAA si possible
        date_eval = row['Date Évaluation']
        if isinstance(date_eval, (pd.Timestamp,)):
            date_eval = date_eval.strftime("%d/%m/%Y")
        else:
            date_eval = str(date_eval)

        repl = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': date_eval
        }
        # appliquer remplacements (paragraphes + cellules)
        for p in doc.paragraphs:
            remplacer_placeholders(p, repl)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl)

        # — Modifications apportées : module unique '1' —

        # Initialisation des compteurs par module (unique '1')
        total_par_module = {'1': 0}
        correct_par_module = {'1': 0}
        questions_par_module = {'1': 0}

        # Compter le nombre de questions pour le module unique
        for q in st.session_state.questions:
            module_key = '1'
            questions_par_module[module_key] = questions_par_module.get(module_key, 0) + 1
            total_par_module[module_key] = questions_par_module[module_key]
            correct_par_module[module_key] = 0

        corr = st.session_state.get('correct_answers', {})
        score_total = 0
        total_questions = questions_par_module['1']  # Total des questions

        # Traiter chaque question pour mélanger/ranger les réponses et compter les bonnes
        for q in st.session_state.questions:
            module_key = '1'
            reps = q['reponses'].copy()
            q_num = q['numero']  # clé pour chercher dans corr

            # Si question "figée", on place la réponse choisie en premier
            if st.session_state.figees.get(q['index'], False):
                bi = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
                cr = reps.pop(bi)
                reps.insert(0, cr)
            else:
                # sinon, on place d'abord la bonne réponse, puis on mélange le reste
                if q['correct_idx'] is not None:
                    cr = reps.pop(q['correct_idx'])
                    reps.insert(0, cr)
                random.shuffle(reps)

            # Écriture des réponses dans le document
            for r in reps:
                idx = r['index']
                if idx < len(doc.paragraphs):
                    box = "☑" if reps.index(r) == 0 else "☐"
                    doc.paragraphs[idx].clear()
                    doc.paragraphs[idx].add_run(f"{r['lettre']} - {r['texte']} {box}")

            # Comptage du score (module '1' unique) et total
            if q_num in corr and reps[0]['lettre'].upper() == corr[q_num]:
                correct_par_module[module_key] += 1
                score_total += 1

        # Préparation des remplacements finaux pour le module unique '1'
        sr = {}
        score_mod = correct_par_module['1']
        sr['{{result_mod1}}'] = str(score_mod)
        sr['{{total_mod1}}'] = str(total_par_module['1'])

        # Score total et total des questions
        sr['{{result_mod_total}}'] = str(score_total)
        sr['{{total_questions}}'] = str(total_questions)

        # Calcul de l'évaluation finale
        resultat_global = calculer_resultat_final(score_total, total_questions)
        sr['{{result_evaluation}}'] = resultat_global

        # Appliquer remplacements finaux (paragraphes + cellules)
        for p in doc.paragraphs:
            remplacer_placeholders(p, sr)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, sr)

        # Remplacer dans les en-têtes et pieds de page
        for section in doc.sections:
            for header in section.header.paragraphs:
                remplacer_placeholders(header, sr)
            for footer in section.footer.paragraphs:
                remplacer_placeholders(footer, sr)

        return doc, score_total, resultat_global, total_questions

    except Exception as e:
        st.error(f"Erreur génération doc : {e}")
        st.error(traceback.format_exc())
        return None, 0, "Erreur", 0

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
                doc_out, sc, re, tot_q = generer_document(row, st.session_state.doc_template)
                if doc_out:
                    recap.append({
                        "Prénom": row["Prénom"],
                        "Nom": row["Nom"],
                        "Réf": row["Référence Session"],
                        "Score": sc,
                        "Total Questions": tot_q,
                        "Pourcentage": f"{(sc/tot_q)*100:.1f}%" if tot_q > 0 else "0%",
                        "Résultat": re
                    })
                    # Sauvegarde du fichier Word dans un BytesIO pour l'ajouter au ZIP
                    bytes_io = io.BytesIO()
                    doc_out.save(bytes_io)
                    fn = f"QCM_{row['Prénom']}_{row['Nom']}.docx"
                    zf.writestr(fn, bytes_io.getvalue())
                prog.progress((i+1)/total)

            if recap:
                df_r = pd.DataFrame(recap)
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_r.to_excel(writer, index=False, sheet_name="Récapitulatif")
                excel_buffer.seek(0)
                zf.writestr("Recapitulatif_QCM.xlsx", excel_buffer.getvalue())

        buf.seek(0)
        st.success("✅ Génération terminée")
        st.download_button(
            "⬇️ Télécharger ZIP", data=buf,
            file_name="QCM_Personnalises.zip",
            mime="application/zip"
        )
        
        # Afficher un aperçu du récapitulatif
        st.subheader("Récapitulatif des résultats")
        st.dataframe(pd.DataFrame(recap))
        
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
