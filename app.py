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
    puis réaffecte paragraph.text.
    """
    texte = paragraph.text
    if not texte:
        return
    for key, value in replacements.items():
        texte = texte.replace(key, value)
    paragraph.text = texte

def detecter_questions(doc):
    """
    Détecte les questions dans le document Word, en associant à chaque question son module.
    Supporte :
      - En-têtes "Module X :" pour définir le contexte de module.
      - Questions numérotées (1.1, 2.3, etc.) se terminant par '?'
      - Questions non numérotées commençant par '-' et finissant par '?'
      - Réponses A–D (on ne requiert plus la coche pour la bonne réponse)
    Retourne une liste de dicts avec clés : index, texte, numero, module, reponses.
    """
    questions = []
    current_question = None
    compteur_non_numerote = 0
    current_module = None

    pattern_module = re.compile(r'^\s*Module\s+(\d+)\s*[:\-]?', re.IGNORECASE)
    pattern_num = re.compile(r'^\s*(\d+(?:\.\d+)*)\s*[-\s–—.]*\s*(.+?)\s*\?$')
    pattern_non_num = re.compile(r'^\s*[-–—]\s*(.+?)\s*\?$')
    # Réponse A–D, on n’exige plus {{checkbox}} ou ☑ ici
    pattern_reponse = re.compile(r'^([A-D])\s*[-\s–—.]+\s*(.+)$', re.IGNORECASE)

    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()\
                    .replace("\u00a0", " ")\
                    .replace("–", "-")\
                    .replace("—", "-")
        if not texte:
            continue

        # 1) Détection d'un en-tête de module
        m_mod = pattern_module.match(texte)
        if m_mod:
            current_module = m_mod.group(1)
            continue

        # 2) Question numérotée
        m_num = pattern_num.match(texte)
        if m_num:
            num = m_num.group(1)
            txt = m_num.group(2)
            libelle = f"{num} - {txt}?"
            current_question = {
                "index": i,
                "texte": libelle,
                "numero": num,
                "module": current_module if current_module else "0",
                "reponses": []
            }
            questions.append(current_question)
            continue

        # 3) Question non-numérotée
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
                "module": current_module if current_module else "0",
                "reponses": []
            }
            questions.append(current_question)
            continue

        # 4) Si on est dans une question en cours, détecter une réponse A–D
        if current_question:
            m_rep = pattern_reponse.match(texte)
            if m_rep:
                lettre = m_rep.group(1).upper()
                rep_txt = m_rep.group(2).strip()
                # On ignore une éventuelle coche dans le texte
                rep_txt = rep_txt.replace("☑", "").replace("☐", "").strip()
                current_question["reponses"].append({
                    "index": i,
                    "lettre": lettre,
                    "texte": rep_txt
                })
                continue

    # Conserver toutes les questions ayant au moins 2 réponses
    valid = []
    for q in questions:
        if len(q["reponses"]) >= 2:
            valid.append(q)
        else:
            st.warning(f"Ignorée : {q['texte']} (moins de 2 propositions)")
    return valid

def parse_correct_answers(f):
    """
    Parse le fichier Excel des réponses correctes et retourne un dict {question_num: lettre_correcte}.
    """
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
    """
    Génère un document QCM pour jedi ligne Excel.
    Calcule dynamiquement les scores par module en se basant sur le fichier de corrections.
    """
    try:
        doc = Document(io.BytesIO(template_bytes))

        # --- 1) Remplacement des placeholders apprenant ---
        repl_apprenant = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(row['Date Évaluation'])
        }
        for p in doc.paragraphs:
            remplacer_placeholders(p, repl_apprenant)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl_apprenant)

        # --- 2) Calculs dynamiques par module ---
        # a) On identifie tous les modules présents
        modules_detectes = set(q['module'] for q in st.session_state.questions)
        # b) Initialiser compteurs
        total_par_module   = {mod: 0 for mod in modules_detectes}
        correct_par_module = {mod: 0 for mod in modules_detectes}
        score_total = 0
        corr = st.session_state.get('correct_answers', {})

        # c) Comptabiliser le total de questions par module
        for q in st.session_state.questions:
            total_par_module[q['module']] += 1

        # d) Pour chaque question, ordonner, écrire et scorer
        for q in st.session_state.questions:
            mod = q['module']
            reps = q['reponses'].copy()
            q_num = q['numero']

            # Si question "figée", conserver la position choisie
            if st.session_state.figees.get(q['index'], False):
                bi = st.session_state.reponses_correctes.get(q['index'], 0)
                cr = reps.pop(bi)
                reps.insert(0, cr)
            else:
                # Sinon, on place d'abord la bonne réponse selon le fichier Excel
                if q_num in corr:
                    lettre_corr = corr[q_num]
                    # Trouver l'index dans reps où lettre == lettre_corr
                    idx_corr = next((i for i, r in enumerate(reps) if r['lettre'] == lettre_corr), None)
                    if idx_corr is not None:
                        r_corr = reps.pop(idx_corr)
                        reps.insert(0, r_corr)
                # Puis on mélange le reste
                random.shuffle(reps[1:])

            # Écriture des réponses dans le document
            for r in reps:
                idx = r['index']
                if idx < len(doc.paragraphs):
                    box = "☑" if reps.index(r) == 0 else "☐"
                    doc.paragraphs[idx].clear()
                    doc.paragraphs[idx].add_run(f"{r['lettre']} - {r['texte']} {box}")

            # Scoring : comparer la lettre en position 0 avec la correction
            if q_num in corr and reps[0]['lettre'] == corr[q_num]:
                correct_par_module[mod] += 1
                score_total += 1

        # --- 3) Remplacement des placeholders {{result_modX}} ---
        repl_scores = {}
        for mod in modules_detectes:
            placeholder = f"{{{{result_mod{mod}}}}}"
            repl_scores[placeholder] = str(correct_par_module[mod])
        # Total général
        repl_scores['{{result_mod_total}}'] = str(score_total)
        # Résultat global
        resultat_global = calculer_resultat_final(score_total, sum(total_par_module.values()))
        repl_scores['{{result_evaluation}}'] = resultat_global

        for p in doc.paragraphs:
            remplacer_placeholders(p, repl_scores)
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl_scores)

        # --- 4) Mise à jour dynamique du tableau des résultats ---
        if doc.tables:
            table = doc.tables[0]
            # Supprimer toutes les lignes sauf l'en-tête (indice 0)
            for row in table.rows[1:]:
                table._tbl.remove(row._tr)
            # Ajouter une ligne par module
            for mod in sorted(modules_detectes, key=lambda x: int(x) if x.isdigit() else float('inf')):
                tot_q = total_par_module[mod]
                score_mod = correct_par_module[mod]
                cells = table.add_row().cells
                cells[0].text = f"Module {mod}"
                cells[1].text = str(score_mod)
                cells[2].text = f"{tot_q} questions"
                cells[3].text = ""
            # Ligne Total
            sum_tot = sum(total_par_module.values())
            cells = table.add_row().cells
            cells[0].text = "Total"
            cells[1].text = str(score_total)
            cells[2].text = f"{sum_tot} questions"
            cells[3].text = ""

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

# Charger Word & détecter questions
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
                    st.write(f"**{idx}. [Module {q['module']}] {q['texte']}**")
                    for r in q['reponses']:
                        st.write(f"- {r['lettre']} : {r['texte']}")
        else:
            st.warning("⚠️ Aucune question détectée. Vérifiez le format.")
    except Exception as e:
        st.error(f"Erreur chargement Word : {e}")
        st.error(traceback.format_exc())

# Charger corrections
if corr_file:
    st.session_state.correct_answers = parse_correct_answers(corr_file)
    st.success(f"✅ {len(st.session_state.correct_answers)} corrections chargées")

# Configuration des questions figées
if st.session_state.questions:
    st.markdown("### Configuration des questions")
    for q in st.session_state.questions:
        with st.expander(f"[Module {q['module']}] {q['texte']}", expanded=False):
            fig = st.checkbox("Figer cette question", key=f"figer_{q['index']}")
            if fig:
                opts = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
                st.session_state.figees[q['index']] = True
                default = 0
                choix = st.selectbox("Bonne réponse", opts, index=default, key=f"bonne_{q['index']}")
                st.session_state.reponses_correctes[q['index']] = opts.index(choix)

# Génération des QCM
if excel_file and word_file and st.session_state.questions and st.button("Générer les QCM"):
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

# Légende résultats
st.markdown("### Légende résultats")
st.info("""
- **Acquis** : ≥ 75 %  
- **En cours d’acquisition** : 50–75 %  
- **Non acquis** : < 50 %
""")
