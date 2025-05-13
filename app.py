# ... (le reste du code reste inchangé jusqu'au file_uploader)

if st.button("4. Générer les fichiers QCM") and excel_file and word_file:
    with tempfile.TemporaryDirectory() as tmpdirname:
        try:
            df = pd.read_excel(excel_file)
            word_path = os.path.join(tmpdirname, "template.docx")
            with open(word_path, "wb") as f:
                f.write(word_file.getbuffer())

            zip_path = os.path.join(tmpdirname, "QCM_generes.zip")
            with ZipFile(zip_path, 'w') as zipf:
                total = len(df)

                for i, row in df.iterrows():
                    # Récupération de toutes les variables
                    prenom = str(row["Prénom"])
                    nom = str(row["Nom"])
                    email = str(row["Email"])
                    ref_session = str(row["Référence Session"])
                    date_eval = str(row["Date Évaluation"])
                    
                    # Nettoyage des noms pour le fichier
                    safe_prenom = re.sub(r'[\\/*?:"<>|]', "_", prenom)
                    safe_nom = re.sub(r'[\\/*?:"<>|]', "_", nom)
                    
                    doc = Document(word_path)

                    # Remplacement de tous les placeholders
                    for para in doc.paragraphs:
                        para.text = (para.text
                                    .replace("{{prenom}}", prenom)
                                    .replace("{{nom}}", nom)
                                    .replace("{{email}}", email)
                                    .replace("{{ref_session}}", ref_session)
                                    .replace("{{date_evaluation}}", date_eval))

                    # Traitement des questions (reste inchangé)
                    j = 0
                    while j < len(doc.paragraphs):
                        if doc.paragraphs[j].text.strip().endswith("?"):
                            if j in figees and j in reponses_correctes:
                                bonne_original = reponses_correctes[j]
                                # Ajout des remplacements pour les bonnes réponses
                                bonne_replaced = (bonne_original
                                                 .replace("{{prenom}}", prenom)
                                                 .replace("{{nom}}", nom)
                                                 .replace("{{email}}", email)
                                                 .replace("{{ref_session}}", ref_session)
                                                 .replace("{{date_evaluation}}", date_eval))
                                figer_reponses(doc.paragraphs, j, bonne_replaced)
                            else:
                                melanger_reponses(doc.paragraphs, j)
                        j += 1

                    # ... (le reste du code reste inchangé)
