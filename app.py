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
    """
    Reconstruit le texte complet du paragraphe, remplace tous les placeholders,
    puis réaffecte paragraph.text pour écraser les runs existants.
    Cela garantit que même si un placeholder est fractionné en plusieurs runs,
    il sera tout de même remplacé.
    """
    texte = paragraph.text
    if not texte:
        return
    # On fait les remplacements successivement sur le texte complet
    for key, value in replacements.items():
        texte = texte.replace(key, value)
        # On gère éventuellement le cas où le placeholder est écrit avec des espaces insécables
        # mais dans notre cas, key = "{{prenom}}" etc. ne contient pas d'espaces, on peut omettre ce passage.
        # Si besoin, on pourrait ajouter :
        # ni = key.replace(" ", "\u00a0")
        # texte = texte.replace(ni, value)
        # ns = key.replace(" ", "")
        # texte = texte.replace(ns, value)
    # Écraser le paragraphe existant et écrire le texte remplacé
    paragraph.text = texte

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
        return "Acquis"
    elif pct >= 50:
        return "En cours d'acquisition"
    else:
        return "Non acquis"

def generer_document(row, template_bytes):
    try:
        doc = Document(io.BytesIO(template_bytes))

        # --- 1) Remplacement des placeholders apprénant ---
        repl_apprenant = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }

        # Appliquer d'abord aux paragraphes
        for p in doc.paragraphs:
            remplacer_placeholders(p, repl_apprenant)
        # Puis aux cellules des tableaux (au cas où des placeholders s'y trouvent aussi)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl_apprenant)

        # --- 2) Initialisation des compteurs par module ---
        total_par_module = {}
        correct_par_module = {}
        # Comptabiliser le nombre de questions par module
        for q in st.session_state.questions:
            # Clé module = partie avant le premier point si numéroté, sinon numéro fictif "NNx"
            if q['numero'].startswith("NN"):
                module_key = q['numero']
            else:
                module_key = q['numero'].split('.')[0]
            total_par_module[module_key] = total_par_module.get(module_key, 0) + 1
            correct_par_module[module_key] = 0

        corr = st.session_state.get('correct_answers', {})
        score_total = 0

        # --- 3) Traitement de chaque question pour mélange / rangement et scoring ---
        for q in st.session_state.questions:
            module_key = q['numero'].split('.')[0] if not q['numero'].startswith("NN") else q['numero']
            reps = q['reponses'].copy()
            q_num = q['numero']

            # Si question figée, on place la réponse choisie en premier
            if st.session_state.figees.get(q['index'], False):
                bi = st.session_state.reponses_correctes.get(q['index'], q['correct_idx'])
                cr = reps.pop(bi)
                reps.insert(0, cr)
            else:
                # Sinon, on place d’abord la bonne réponse, puis on mélange le reste
                if q['correct_idx'] is not None:
                    cr = reps.pop(q['correct_idx'])
                    reps.insert(0, cr)
                random.shuffle(reps)

            # Écrire les réponses dans le document (on réécrit chaque paragraphe original)
            for r in reps:
                idx = r['index']
                if idx < len(doc.paragraphs):
                    box = "☑" if reps.index(r) == 0 else "☐"
                    doc.paragraphs[idx].clear()
                    doc.paragraphs[idx].add_run(f"{r['lettre']} - {r['texte']} {box}")

            # Comptage du score par module et total
            if q_num in corr and reps[0]['lettre'].upper() == corr[q_num]:
                correct_par_module[module_key] += 1
                score_total += 1

        # --- 4) Préparer les remplacements finaux containers de résultats ---
        repl_resultats = {}
        # Pour chaque module, remplacer {{result_modX}} par le score brut
        for module_key, tot in total_par_module.items():
            score_mod = correct_par_module.get(module_key, 0)
            repl_resultats[f'{{{{result_mod{module_key}}}}}'] = str(score_mod)
            # Si vous souhaitez aussi un libellé par module, vous pourriez faire :
            # repl_resultats[f'{{{{result_mod{module_key}_eval}}}}'] = \
            #     calculer_resultat_final(score_mod, tot)

        # Remplacement du score total et du résultat global
        repl_resultats['{{result_mod_total}}'] = str(score_total)
        resultat_global = calculer_resultat_final(score_total, sum(total_par_module.values()))
        repl_resultats['{{result_evaluation}}'] = resultat_global

        # Appliquer ces remplacements finaux (paragraphes + tableaux)
        for p in doc.paragraphs:
            remplacer_placeholders(p, repl_resultats)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl_resultats)

        return doc, score_total, resultat_global

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

# Initialisation de la session
for key in ('questions','figees','reponses_correctes'):
    if key not in st.session_state:
        st.session_state[key] = [] if key=='questions' else {}
if 'current_template' not in st.session_state:
    st.session_state.current_template = None
if 'doc_template' not in st.session_state:
    st.session_state.doc_template = None

# Charger le document Word & détecter les questions
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

# Charger le fichier de corrections, s’il existe
if corr_file:
    st.session_state.correct_answers = parse_correct_answers(corr_file)
    st.success(f"✅ {len(st.session_state.correct_answers)} corrections chargées")

# Configuration des questions figées
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

# Génération des QCM
if excel_file and st.session_state.questions and st.button("Générer les QCM"):
    try:
        df = pd.read_excel(excel_file)
        need = ['Prénom','Nom','Email','Référence Session','Date Évaluation']
        miss = [c for c in need if c not in df.columns]
        if miss:
            st.error(f"Colonnes manquantes : {miss}")
            st.stop()

        buf   = io.BytesIO()
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
                    fn  = f"QCM_{row['Prénom']}_{row['Nom']}.docx"
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
                    doc_out.save(tmp.name)
                    zf.write(tmp.name, fn)
                prog.progress((i+1)/total)

            if recap:
                df_r  = pd.DataFrame(recap)
                tmp2  = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
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

# Légende des résultats
st.markdown("### Légende résultats")
st.info("""
- **Acquis** : ≥ 75%  
- **En cours d'acquisition** : 50–75%  
- **Non acquis** : < 50%
""")
