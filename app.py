import streamlit as st
import pandas as pd
from docx import Document
import random
import io
import traceback
from zipfile import ZipFile
import re

st.set_page_config(page_title="Générateur de QCM", layout="centered")
st.title("Générateur de QCM personnalisés")

# — Fonctions utilitaires —
def remplacer_placeholders(paragraph, replacements):
    """Remplace les placeholders en gérant toutes les variantes d'espaces"""
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
            key.replace("{", "").replace("}", "")  # checkbox seul
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
    """
    Détecte les questions avec gestion des formats complexes
    et espaces dans les numéros de question
    """
    questions = []
    current_question = None
    
    # Regex améliorée pour gérer les formats spécifiques du fichier
    question_pattern = re.compile(
        r'^\s*'                          # Espaces initiaux
        r'(\d+(?:[\s\.]*\d+)*)'          # Numéro de question avec espaces et points
        r'[\s\-–—.]+'                    # Séparateur (tiret ou point)
        r'(.+?)'                         # Texte de la question
        r'\s*\?$'                        # Point d'interrogation en fin de ligne
    )
    
    # Regex pour gérer les réponses avec ou sans espaces autour de {{checkbox}}
    reponse_pattern = re.compile(
        r'^([A-D])'                          # Lettre de réponse (A-D)
        r'[\s\-–—.]+'                        # Séparateurs (espaces, tirets, points)
        r'(.+?)'                             # Texte de la réponse
        r'(\{\{\s*checkbox\s*\}\})?'         # Bonne réponse avec espaces optionnels
        r'$',                                # Fin de ligne
        re.IGNORECASE
    )
    
    for i, para in enumerate(doc.paragraphs):
        texte = para.text.strip()
        
        # Détection des questions
        match = question_pattern.match(texte)
        if match:
            # Nettoyage du numéro de question
            numero_brut = match.group(1)
            numero = re.sub(r'[\s\.]+', '.', numero_brut).strip()
            numero = re.sub(r'\.+$', '', numero)  # Supprime les points en trop
            
            texte_question = match.group(2).strip()
            texte_complet = f"{numero} - {texte_question}?"
            
            current_question = {
                "index": i,
                "texte": texte_complet,
                "numero": numero,
                "reponses": [],
                "correct_idx": None,
                "original_text": texte
            }
            questions.append(current_question)
        
        # Détection des réponses
        elif current_question:
            match_reponse = reponse_pattern.match(texte)
            if match_reponse:
                lettre = match_reponse.group(1).upper()
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
    valides = [q for q in questions if len(q["reponses"]) >= 2 and q["correct_idx"] is not None]
    
    # Affichage des questions ignorées
    for q in questions:
        if q not in valides:
            st.warning(f"Ignorée : {q['texte']} (moins de 2 réponses ou pas de {{checkbox}})")
    
    return valides

def parse_correct_answers(f):
    """
    Lit un fichier Excel avec les réponses correctes
    et nettoie les numéros de question pour correspondre au format détecté
    """
    if f is None:
        return {}
    
    try:
        df = pd.read_excel(f)
        df = df.dropna(subset=['Numéro de la question', 'Réponse correcte'])
        
        # Nettoyage des numéros de question
        df['Numéro de la question'] = df['Numéro de la question'].astype(str).str.strip()
        df['Numéro de la question'] = df['Numéro de la question'].str.replace(r'\s+', '.', regex=True)
        df['Numéro de la question'] = df['Numéro de la question'].str.replace(r'\.+$', '', regex=True)
        
        df['Réponse correcte'] = df['Réponse correcte'].astype(str).str.strip().str.upper()
        
        return dict(zip(df['Numéro de la question'], df['Réponse correcte']))
    except Exception as e:
        st.error(f"Erreur lecture corrections : {e}")
        return {}

def calculer_resultat_final(score, total_q):
    """
    Calcule le résultat final en fonction du pourcentage
    """
    if total_q <= 0:
        return "Non acquis"
    
    pct = (score / total_q) * 100
    
    if pct >= 75:
        return "Acquis"
    elif pct >= 50:
        return "En cours d’acquisition"
    else:
        return "Non acquis"

def generer_document(row, template_bytes):
    """
    Génère un document Word personnalisé pour un apprenant
    """
    try:
        doc = Document(io.BytesIO(template_bytes))
        
        # Remplacement des placeholders d'apprenant
        date_eval = row['Date Évaluation']
        if isinstance(date_eval, (pd.Timestamp,)):
            date_eval = date_eval.strftime("%d/%m/%Y")
        
        repl_apprenant = {
            '{{prenom}}': str(row['Prénom']),
            '{{nom}}': str(row['Nom']),
            '{{email}}': str(row['Email']),
            '{{ref_session}}': str(row['Référence Session']),
            '{{date_evaluation}}': str(date_eval)
        }
        
        # Remplacement dans les paragraphes
        for p in doc.paragraphs:
            remplacer_placeholders(p, repl_apprenant)
        
        # Remplacement dans les tableaux
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, repl_apprenant)
        
        # Préparation du comptage
        questions = st.session_state['questions']
        total_questions = len(questions)
        score_total = 0
        
        # Dictionnaire pour regrouper par module
        modules = {}
        
        # Traitement de chaque question
        for q in questions:
            # Nettoyage du numéro de question
            numero = q['numero']
            module_key = numero.split('.')[0] if '.' in numero else numero
            
            # Récupération des réponses
            reponses = q['reponses'].copy()
            
            # Mélange des réponses si non figées
            if st.session_state['figees'].get(q['index'], False):
                idx = st.session_state['reponses_correctes'].get(q['index'], q['correct_idx'])
                bonne_rep = reponses.pop(idx)
                reponses.insert(0, bonne_rep)
            else:
                if q['correct_idx'] is not None:
                    bonne_rep = reponses.pop(q['correct_idx'])
                    reponses.insert(0, bonne_rep)
                random.shuffle(reponses)
            
            # Mise à jour du document avec les réponses
            for rep in reponses:
                idx_para = rep['index']
                if idx_para < len(doc.paragraphs):
                    box = "☑" if reponses.index(rep) == 0 else "☐"
                    doc.paragraphs[idx_para].clear()
                    doc.paragraphs[idx_para].add_run(f"{rep['lettre']} - {rep['texte']} {box}")
            
            # Vérification de la réponse
            numero_clean = re.sub(r'[^\d.]', '', numero.split('-')[0].strip())
            
            if numero_clean in st.session_state.correct_answers:
                generated_answer = reponses[0]['lettre'].upper()
                expected_answer = st.session_state.correct_answers[numero_clean].upper()
                
                if generated_answer == expected_answer:
                    score_total += 1
                    modules[module_key] = modules.get(module_key, 0) + 1
        
        # Préparation des remplacements de score
        sr = {
            '{{result_mod_total}}': str(score_total),
            '{{total_questions}}': str(total_questions),
            '{{result_evaluation}}': calculer_resultat_final(score_total, total_questions)
        }
        
        # Ajout des scores par module
        for module_key, score in modules.items():
            sr[f'{{{{result_mod{module_key}}}}}'] = str(score)
            sr[f'{{{{total_mod{module_key}}}}}'] = str(len([q for q in questions if q['numero'].startswith(module_key)]))
        
        # Remplacement des placeholders de score
        for p in doc.paragraphs:
            remplacer_placeholders(p, sr)
        
        for tbl in doc.tables:
            for r in tbl.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        remplacer_placeholders(p, sr)
        
        return doc, score_total, sr['{{result_evaluation}}'], total_questions
    
    except Exception as e:
        st.error(f"Erreur génération doc : {e}")
        st.error(traceback.format_exc())
        return None, 0, "Erreur", 0

# — Interface Streamlit —
with st.expander("Étape 1 : Importation des fichiers", expanded=True):
    excel_file = st.file_uploader("Fichier Excel (colonnes: Prénom, Nom, Email, Référence Session, Date Évaluation)", type="xlsx")
    word_file = st.file_uploader("Modèle Word (.docx)", type="docx")
    corr_file = st.file_uploader("Réponses correctes (Excel .xlsx)", type="xlsx")

# Initialisation du session_state
if 'questions' not in st.session_state:
    st.session_state['questions'] = []
if 'figees' not in st.session_state:
    st.session_state['figees'] = {}
if 'reponses_correctes' not in st.session_state:
    st.session_state['reponses_correctes'] = {}
if 'current_template' not in st.session_state:
    st.session_state['current_template'] = None
if 'correct_answers' not in st.session_state:
    st.session_state['correct_answers'] = {}

# Chargement du modèle Word
if word_file and st.session_state['current_template'] != word_file.name:
    try:
        data = word_file.getvalue()
        doc_sample = Document(io.BytesIO(data))
        qs = detecter_questions(doc_sample)
        st.session_state['questions'] = qs
        st.session_state['current_template'] = word_file.name
        st.session_state['figees'] = {}
        st.session_state['reponses_correctes'] = {}
        
        if qs:
            st.success(f"✅ {len(qs)} questions détectées")
            with st.expander("🔍 Questions détectées", expanded=True):
                for idx, q in enumerate(qs, 1):
                    st.write(f"**{idx}. {q['texte']}**")
                    for j, r in enumerate(q['reponses']):
                        mark = "✅" if j == q['correct_idx'] else "☐"
                        st.write(f"{mark} {r['lettre']} - {r['texte']}")
        else:
            st.warning("⚠️ Aucune question détectée. Vérifiez le format du Word.")
    except Exception as e:
        st.error(f"Erreur chargement Word : {e}")
        st.error(traceback.format_exc())

# Chargement des réponses correctes
if corr_file:
    ca = parse_correct_answers(corr_file)
    st.session_state['correct_answers'] = ca
    st.success(f"✅ {len(ca)} corrections chargées")

# Configuration des questions figées
if st.session_state['questions']:
    st.markdown("### Configuration des questions")
    for q in st.session_state['questions']:
        q_id = q['index']
        q_num = q['numero']
        key_base = f"{q_id}_{st.session_state['current_template']}"
        fig = st.checkbox(f"Figer question {q_num}", key=f"figer_{key_base}")
        
        if fig:
            options = [f"{r['lettre']} - {r['texte']}" for r in q['reponses']]
            default_idx = q['correct_idx'] if q['correct_idx'] is not None else 0
            choix = st.selectbox(
                f"Choix pour {q_num}",
                options=options,
                index=default_idx,
                key=f"bonne_{key_base}"
            )
            st.session_state['figees'][q_id] = True
            st.session_state['reponses_correctes'][q_id] = options.index(choix)
        else:
            st.session_state['figees'].pop(q_id, None)
            st.session_state['reponses_correctes'].pop(q_id, None)

# Génération des QCM
if excel_file and word_file and st.session_state['questions']:
    if st.button("Générer les QCM"):
        try:
            df = pd.read_excel(excel_file)
            required_cols = ['Prénom', 'Nom', 'Email', 'Référence Session', 'Date Évaluation']
            missing = [c for c in required_cols if c not in df.columns]
            
            if missing:
                st.error(f"Colonnes manquantes dans l'Excel : {missing}")
                st.stop()
            
            buf = io.BytesIO()
            recap = []
            
            with ZipFile(buf, 'w') as zf:
                progress = st.progress(0)
                total = len(df)
                
                for i, row in df.iterrows():
                    doc_out, sc, res, tot_q = generer_document(row, word_file.getvalue())
                    
                    if doc_out:
                        recap.append({
                            "Prénom": row["Prénom"],
                            "Nom": row["Nom"],
                            "Réf": row["Référence Session"],
                            "Score": sc,
                            "Total Questions": tot_q,
                            "Pourcentage": f"{(sc/tot_q)*100:.1f}%" if tot_q > 0 else "0%",
                            "Résultat": res
                        })
                        
                        bytes_io = io.BytesIO()
                        doc_out.save(bytes_io)
                        fn = f"QCM_{row['Prénom']}_{row['Nom']}.docx"
                        zf.writestr(fn, bytes_io.getvalue())
                    
                    progress.progress((i + 1) / total)
                
                if recap:
                    df_r = pd.DataFrame(recap)
                    excel_buf = io.BytesIO()
                    
                    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                        df_r.to_excel(writer, index=False, sheet_name="Récapitulatif")
                    
                    excel_buf.seek(0)
                    zf.writestr("Recapitulatif_QCM.xlsx", excel_buf.getvalue())
            
            buf.seek(0)
            st.success("✅ Génération terminée")
            
            st.download_button(
                "⬇️ Télécharger l’archive ZIP",
                data=buf,
                file_name="QCM_Personnalises.zip",
                mime="application/zip"
            )
            
            st.subheader("Récapitulatif des résultats")
            st.dataframe(pd.DataFrame(recap))
        
        except Exception as e:
            st.error(f"ERREUR critique : {e}")
            st.error(traceback.format_exc())

# Légende des résultats
st.markdown("### Légende résultats")
st.info("""
- **Acquis** : ≥ 75%  
- **En cours d’acquisition** : 50–75%  
- **Non acquis** : < 50%
""")
