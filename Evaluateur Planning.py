import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import sqlite3
from dateutil import parser
import re  # Pour les expressions régulières
from fpdf import FPDF
from tkinter import PhotoImage
from PIL import Image, ImageTk

# Variable globale pour stocker le chemin du fichier sélectionné
file_path = None

# Connexion à la base de données SQLite existante
conn = sqlite3.connect('BDD_GDP.sqlite3')
cursor = conn.cursor()

# Liste des colonnes attendues
colonnes_attendues = [
    "WBS", "Nom", "Durée_prévue", "Début", "Fin", 
    "Prédécesseurs", "Successeurs", "Pourcentage_achevé", 
    "Type_de_contrainte", "Marge_totale", "N°"
]

# Fonction pour convertir les mois français en anglais
def remplacer_mois_fr_en(date_str):
    mois_fr = {
        "Janvier": "January",
        "Février": "February",
        "Mars": "March",
        "Avril": "April",
        "Mai": "May",
        "Juin": "June",
        "Juillet": "July",
        "Août": "August",
        "Septembre": "September",
        "Octobre": "October",
        "Novembre": "November",
        "Décembre": "December"
    }
    
    for fr, en in mois_fr.items():
        if fr in date_str:
            date_str = date_str.replace(fr, en)
    
    return date_str

# Fonction pour convertir les dates en utilisant dateutil
def convertir_date(date_str):
    if not date_str or date_str.strip() == "":
        return None
    
    # Remplacer les mois en français par leurs équivalents en anglais
    date_str_en = remplacer_mois_fr_en(date_str).strip()

    try:
        # Utiliser dateutil pour parser la date
        date_obj = parser.parse(date_str_en)
        return date_obj.isoformat()
    except Exception as e:
        return None

# Fonction pour insérer le projet dans la BDD
def inserer_projet(nom_projet):
    try:
        cursor.execute("INSERT INTO Projet (Nom_projet) VALUES (?)", (nom_projet,))
        conn.commit()
        return cursor.lastrowid  # Récupérer l'ID du projet inséré
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'insertion du projet : {str(e)}")
        return None

# Fonction pour mettre à jour les dates de début et de fin du projet
def mettre_a_jour_dates_projet(projet_id, date_debut, date_fin):
    try:
        cursor.execute("""
            UPDATE Projet 
            SET Date_debut = ?, Date_fin = ?
            WHERE id = ?
        """, (date_debut, date_fin, projet_id))
        conn.commit()
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la mise à jour des dates du projet : {str(e)}")

# Fonction pour obtenir l'ID maximum actuel dans la table Planning
def obtenir_max_id_planning():
    cursor.execute("SELECT MAX(id) FROM Planning")
    result = cursor.fetchone()[0]
    return result if result is not None else 0

# Fonction pour décaler les IDs des prédécesseurs et successeurs
def decaler_ids_predecesseurs_successeurs(valeur, decalage):
    if pd.isna(valeur):
        return valeur
    
    # Séparer les différentes parties basées sur le ";" et traiter chaque partie individuellement
    parties = valeur.split(";")
    
    nouvelles_parties = []
    for partie in parties:
        # Extraire l'ID numérique au début de la chaîne (avant "FD", "DD", etc.)
        match = re.match(r"(\d+)(.*)", partie)
        if match:
            id_num = int(match.group(1))  # ID de la tâche
            suffixe = match.group(2)  # Suffixe comme "FD+5j", "DD", etc.
            id_num += decalage  # Décaler l'ID
            nouvelles_parties.append(f"{id_num}{suffixe}")  # Recomposer la chaîne avec l'ID décalé et le suffixe
        else:
            nouvelles_parties.append(partie)  # Si pas d'ID trouvé, laisser la partie inchangée

    return ";".join(nouvelles_parties)

# Fonction pour importer le fichier Excel
def importer_fichier():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            colonnes_presentes = df.columns.tolist()

            # Vérification des colonnes manquantes
            colonnes_manquantes = [col for col in colonnes_attendues if col not in colonnes_presentes]

            if colonnes_manquantes:
                messagebox.showerror(
                    "Erreur de colonnes",
                    f"Les colonnes suivantes sont manquantes : {', '.join(colonnes_manquantes)}.\n"
                    "Merci de renommer le fichier avec les colonnes suivantes :\n"
                    "WBS, Nom, Durée_prévue, Début, Fin, Prédécesseurs, "
                    "Successeurs, Pourcentage_achevé, Type_de_contrainte, Marge_totale, N°"
                )
                btn_traiter.config(state="disabled")
            else:
                messagebox.showinfo("Fichier Sélectionné", f"Fichier sélectionné : {file_path}")
                btn_traiter.config(state="normal")

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors de la lecture du fichier : {str(e)}")
            btn_traiter.config(state="disabled")
    else:
        messagebox.showwarning("Avertissement", "Aucun fichier sélectionné.")
        btn_traiter.config(state="disabled")

# Fonction pour traiter le fichier et ajuster les IDs du planning
def traiter_fichier():
    global entry_nom_projet
    nom_projet = entry_nom_projet.get().strip()  # Récupérer le nom du projet entré par l'utilisateur

    if not nom_projet:
        messagebox.showerror("Erreur", "Veuillez entrer un nom de projet.")
        return

    if file_path:
        try:
            # Insérer le projet et récupérer son ID
            projet_id = inserer_projet(nom_projet)
            if not projet_id:
                return  # Si l'insertion du projet échoue, on arrête le processus

            df = pd.read_excel(file_path, engine='openpyxl')

            # Nombre total de lignes pour la barre de progression
            total_lignes = len(df)

            # Initialiser la barre de progression
            progress_bar['value'] = 0
            root.update_idletasks()

            # Conversion des dates au format ISO 8601
            df['Début'] = df['Début'].apply(lambda date: convertir_date(date))
            df['Fin'] = df['Fin'].apply(lambda date: convertir_date(date))

            # Mettre à jour la barre de progression après conversion des dates
            progress_bar['value'] = 50
            root.update_idletasks()

            # Obtenir l'ID maximum actuel dans la table Planning
            max_id = obtenir_max_id_planning()

            # Ajouter un décalage aux IDs dans la colonne N° pour éviter les conflits
            df['N°'] = df['N°'] + max_id + 1

            # Décaler les IDs des prédécesseurs et successeurs
            df['Prédécesseurs'] = df['Prédécesseurs'].apply(lambda val: decaler_ids_predecesseurs_successeurs(val, max_id + 1))
            df['Successeurs'] = df['Successeurs'].apply(lambda val: decaler_ids_predecesseurs_successeurs(val, max_id + 1))

            # Sélectionner les colonnes nécessaires pour la BDD
            df_selection = df[['WBS', 'Nom', 'Durée_prévue', 'Début', 'Fin', 
                               'Pourcentage_achevé', 'Type_de_contrainte', 'Marge_totale', 
                               'Prédécesseurs', 'Successeurs', 'N°']].copy()

            # Ajouter la colonne Code_projet avec l'ID du projet
            df_selection['Code_projet'] = projet_id

            # Renommer les colonnes pour correspondre à la BDD
            df_selection.rename(columns={
                'WBS': 'Niveau_WBS',
                'Nom': 'Nom',
                'Durée_prévue': 'duree',
                'Début': 'date_debut',
                'Fin': 'date_de_fin',
                'Pourcentage_achevé': 'Avancement',
                'Type_de_contrainte': 'type_contrainte',
                'Marge_totale': 'marge_totale',
                'Prédécesseurs': 'predecesseurs',
                'Successeurs': 'successeurs',
                'N°': 'id'
            }, inplace=True)

            # Insérer les données dans la BDD
            df_selection.to_sql('Planning', conn, if_exists='append', index=False)

            # Récupérer la date de début la plus ancienne et la date de fin la plus récente
            date_debut_min = df_selection['date_debut'].min()
            date_fin_max = df_selection['date_de_fin'].max()

            # Mettre à jour les dates du projet dans la table Projet
            mettre_a_jour_dates_projet(projet_id, date_debut_min, date_fin_max)

            # Mettre à jour la barre de progression après insertion des données
            progress_bar['value'] = 100
            root.update_idletasks()

            # Passer à la nouvelle fenêtre pour sélectionner un rapport
            afficher_page_evaluation()

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors du traitement du fichier : {str(e)}")
            progress_bar['value'] = 0
            root.update_idletasks()
    else:
        messagebox.showwarning("Avertissement", "Aucun fichier à traiter.")
        progress_bar['value'] = 0
        root.update_idletasks()

def generer_pdf(resultats, commentaire, nom_projet):
    pdf = FPDF()
    pdf.add_page()

    # Titre du document avec le nom du projet
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(200, 10, txt=f"Rapport d'évaluation du planning projet : {nom_projet}", ln=True, align='C')

    # Ajouter une section pour le commentaire
    pdf.ln(10)  # Saut de ligne
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(200, 10, txt="Commentaire :", ln=True)
    pdf.set_font('Arial', '', 12)
    pdf.multi_cell(0, 10, txt=commentaire)

    # Saut de ligne avant la table des résultats
    pdf.ln(10)

    # Ajouter les résultats de l'évaluation
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(60, 10, 'Critère', 1)
    pdf.cell(40, 10, 'Occurrences', 1)
    pdf.cell(30, 10, 'Total', 1)
    pdf.cell(30, 10, 'Ratio (%)', 1)
    pdf.cell(30, 10, 'Indicateur', 1)
    pdf.ln()

    pdf.set_font('Arial', '', 12)
    
    # Déterminer la largeur des colonnes
    largeur_colonne_critere = 60
    largeur_colonne_occurences = 40
    largeur_colonne_total = 30
    largeur_colonne_ratio = 30
    largeur_colonne_indicateur = 30
    
    for resultat in resultats:
        # Vérifier si on a suffisamment d'espace pour ajouter une nouvelle ligne, sinon sauter à la page suivante
        if pdf.get_y() > 260:  # Ajustez cette valeur en fonction de vos besoins de marges
            pdf.add_page()

        # Nettoyer les espaces superflus dans les critères
        critere_texte = " ".join(resultat["critere"].split())  # Supprimer les espaces multiples
        
        # Calculer la hauteur de la ligne la plus haute dans chaque rangée
        critere_lines = pdf.get_string_width(critere_texte) / largeur_colonne_critere
        critere_lines = max(1, round(critere_lines) + 1)
        
        occurences_lines = 1
        total_lines = 1
        ratio_lines = 1
        indicateur_lines = 1

        # La hauteur maximale déterminée par le critère le plus grand en nombre de lignes
        max_lines = max(critere_lines, occurences_lines, total_lines, ratio_lines, indicateur_lines)
        line_height = 6  # hauteur d'une ligne
        
        total_height = line_height * max_lines

        # Critère (multicell pour gérer les longs textes)
        y_before = pdf.get_y()  # Sauvegarder la position Y avant d'écrire
        pdf.multi_cell(largeur_colonne_critere, line_height, critere_texte, 1)
        y_after = pdf.get_y()  # Position Y après avoir écrit

        # Revenons à la position initiale pour aligner les autres cellules
        pdf.set_xy(pdf.get_x() + largeur_colonne_critere, y_before)

        # Occurrences, Total et Ratio
        pdf.cell(largeur_colonne_occurences, total_height, str(resultat["nombre_occurences"]), 1)
        pdf.cell(largeur_colonne_total, total_height, str(resultat["nombre_total"]), 1)
        pdf.cell(largeur_colonne_ratio, total_height, str(round(resultat["ratio"], 2)), 1)

        # Si un indicateur n'existe pas, afficher "N/A"
        if "indicateur" in resultat and resultat["indicateur"] is not None:
            # Appliquer la couleur pour l'indicateur
            if resultat["indicateur"] == 1:
                pdf.set_fill_color(255, 0, 0)  # Rouge
            elif resultat["indicateur"] == 3:
                pdf.set_fill_color(255, 165, 0)  # Orange
            elif resultat["indicateur"] == 5:
                pdf.set_fill_color(0, 255, 0)  # Vert
            else:
                pdf.set_fill_color(255, 255, 255)  # Blanc (par défaut)
            pdf.cell(largeur_colonne_indicateur, total_height, str(resultat["indicateur"]), 1, fill=True)
        else:
            # Si l'indicateur n'est pas applicable, afficher "N/A"
            pdf.cell(largeur_colonne_indicateur, total_height, "N/A", 1, fill=True)

        # Revenir à la ligne suivante
        pdf.ln(total_height)

    # Sauvegarder le fichier PDF
    pdf.output("rapport_evaluation.pdf")
    messagebox.showinfo("Succès", "Le rapport a été généré avec succès en PDF.")



def generer_rapport_pdf(projet_id, commentaire, nom_projet):
    # Récupérer les résultats de l'évaluation actuelle
    resultats = evaluer_projet(projet_id)
    
    # Appelle la fonction de génération de PDF avec le commentaire, les résultats et le nom du projet
    generer_pdf(resultats, commentaire, nom_projet)

# Fonction pour afficher la nouvelle page de sélection des rapports
def afficher_page_evaluation():
    # Cacher la fenêtre d'importation
    for widget in root.winfo_children():
        widget.pack_forget()

    # Frame principale pour tout l'écran
    frame_evaluation = ttk.Frame(root, padding="10 10 10 10")
    frame_evaluation.pack(expand=True, fill='both')

    # Frame pour sélectionner le projet
    selection_frame = ttk.Frame(frame_evaluation)
    selection_frame.pack(pady=10)

    tk.Label(selection_frame, text="Sélectionner un planning pour évaluation :").pack(side="left", padx=10)

    # Récupérer les projets depuis la base de données
    projets = cursor.execute("SELECT id, Nom_projet FROM Projet").fetchall()
    projets_dict = {nom: id for id, nom in projets}

    # Combobox pour sélectionner un projet
    combo_projets = ttk.Combobox(selection_frame, values=list(projets_dict.keys()), width=40)
    combo_projets.pack(side="left", padx=10)

    # Bouton pour évaluer le projet à droite de la sélection
    btn_evaluer = ttk.Button(selection_frame, text="Évaluer le projet", command=lambda: afficher_evaluation(projets_dict, combo_projets, tree))
    btn_evaluer.pack(side="left", padx=10)

    # Tableau pour afficher les résultats de l'évaluation (augmenter la hauteur et largeur)
    tree = ttk.Treeview(frame_evaluation, columns=("Critère", "Nombre d'occurrences", "Nombre total", "Ratio (%)", "Seuil OK", "Indicateur"), show="headings", height=15)  # Hauteur augmentée
    tree.heading("Critère", text="Qualité planning détaillé")
    tree.heading("Nombre d'occurrences", text="Nombre d'occurrence", anchor="center")
    tree.heading("Nombre total", text="Nombre total", anchor="center")
    tree.heading("Ratio (%)", text="Ratio (%)", anchor="center")
    tree.heading("Seuil OK", text="Seuil OK", anchor="center")
    tree.heading("Indicateur", text="Indicateurs", anchor="center")

    # Augmenter la largeur des colonnes pour rendre le tableau plus lisible
    tree.column("Critère", width=300)  # Augmentation de la largeur
    tree.column("Nombre d'occurrences", width=150, anchor="center")
    tree.column("Nombre total", width=150, anchor="center")
    tree.column("Ratio (%)", width=100, anchor="center")
    tree.column("Seuil OK", width=150, anchor="center")
    tree.column("Indicateur", width=150, anchor="center")

    tree.pack(pady=10, padx=10, fill='both', expand=True)

    # Ajouter les tags pour définir les couleurs dans le Treeview
    tree.tag_configure('rouge', background='#FFCCCC')  # Rouge
    tree.tag_configure('orange', background='#FFD580')  # Orange
    tree.tag_configure('vert', background='#C6E9C6')    # Vert

    # Ajouter un champ pour entrer un commentaire
    label_commentaire = ttk.Label(frame_evaluation, text="Ajouter un commentaire :", font=("Arial", 12))
    label_commentaire.pack(pady=10)
    champ_commentaire = tk.Text(frame_evaluation, height=5, width=80, font=("Arial", 12))
    champ_commentaire.pack(pady=10)

    # Frame pour les boutons du bas
    frame_boutons = ttk.Frame(frame_evaluation)
    frame_boutons.pack(pady=10)

    # Bouton pour générer le PDF
    btn_generer_pdf = ttk.Button(frame_boutons, text="Générer le rapport en PDF", command=lambda: generer_rapport_pdf(projets_dict[combo_projets.get()], champ_commentaire.get("1.0", tk.END), combo_projets.get()))
    btn_generer_pdf.pack(side="left", padx=10)

    # Bouton pour revenir à l'importation des fichiers
    btn_retour = ttk.Button(frame_boutons, text="Retour à l'importation", command=revenir_page_import)
    btn_retour.pack(side="left", padx=10)


    
def revenir_page_import():
    # Cacher la fenêtre d'évaluation
    for widget in root.winfo_children():
        widget.pack_forget()

    # Réafficher la fenêtre d'importation
    main_frame.pack(expand=True, fill='both')  # Repack avec les bons paramètres
    entry_nom_projet.delete(0, tk.END)  # Vider le champ de texte pour le nom du projet
    btn_traiter.config(state="disabled")  # Désactiver le bouton "Traiter et envoyer le fichier"
    progress_bar['value'] = 0  # Réinitialiser la barre de progression
    centrer_fenetre(root)  # Réinitialiser la géométrie de la fenêtre
    root.update_idletasks()  # Mettre à jour l'interface pour refléter les changements




# Fonction pour identifier les tâches récapitulatives et non récapitulatives
def identifier_taches_recap(taches):
    """
    Identifie les tâches récapitulatives et non récapitulatives en utilisant le WBS.
    Une tâche est récapitulative si elle a au moins une sous-tâche dont le WBS commence par son propre WBS suivi d'un point,
    et elle est immédiatement suivie par cette sous-tâche dans le fichier.
    """
    taches_recap = []
    taches_non_recap = []

    # Extraire uniquement les colonnes nécessaires (ID de la tâche et WBS)
    taches = [(tache[0], tache[1]) for tache in taches]  # Assurez-vous que la première colonne est l'ID et la seconde est le WBS.

    # Tri des tâches en fonction du WBS pour garantir la continuité des sous-tâches
    taches = sorted(taches, key=lambda x: list(map(int, x[1].split('.'))))

    for i, (id_tache, wbs) in enumerate(taches):
        # Les tâches avec WBS "0" sont toujours récapitulatives
        if wbs == "0":
            taches_recap.append(id_tache)
            continue

        # Vérifier si cette tâche est récapitulative
        est_recap = False
        for j in range(i + 1, len(taches)):
            autre_id_tache, autre_wbs = taches[j]
            # Si le WBS de l'autre tâche commence par le WBS de la tâche courante, c'est une sous-tâche
            if autre_wbs.startswith(f"{wbs}."):
                est_recap = True
                break
            # Si on rencontre une tâche qui ne fait plus partie de la hiérarchie, on arrête de chercher
            elif autre_wbs.split('.')[0] != wbs.split('.')[0]:
                break

        if est_recap:
            taches_recap.append(id_tache)
        else:
            taches_non_recap.append(id_tache)

    return taches_recap, taches_non_recap


# Fonction pour calculer les jours ouvrés entre deux dates
def calculer_jours_ouvres(date_debut, date_fin):
    # Convertir les dates en objets datetime
    date_debut = pd.to_datetime(date_debut)
    date_fin = pd.to_datetime(date_fin)
    
    # Créer une série de jours ouvrés entre les deux dates
    jours_ouvres = pd.bdate_range(start=date_debut, end=date_fin)
    
    # Retourner le nombre de jours ouvrés
    return len(jours_ouvres)


# Fonction pour évaluer les critères basés sur les tâches récapitulatives
def evaluer_projet(projet_id):
    # Récupérer toutes les tâches du projet, y compris la colonne duree, date_debut et date_fin
    cursor.execute("""
        SELECT id, Niveau_WBS, successeurs, predecesseurs, type_contrainte, duree, date_debut, date_de_fin, marge_totale
        FROM Planning 
        WHERE Code_projet = ?
    """, (projet_id,))
    taches = cursor.fetchall()

    # Identifier les tâches récapitulatives et non récapitulatives
    taches_recap, taches_non_recap = identifier_taches_recap(taches)

    # Filtrer les tâches non récapitulatives
    taches_non_recap_non_jalon = [t for t in taches if t[0] in taches_non_recap and not t[5].startswith("0")]
    taches_non_recap_jalon = [t for t in taches if t[0] in taches_non_recap and t[5].startswith("0")]  # Tâches non récap et jalons

    # Nombre total de tâches non récapitulatives qui ne sont pas des jalons
    nombre_total_taches_non_recap_non_jalon = len(taches_non_recap_non_jalon)
    # Nombre total de tâches non récapitulatives qui sont des jalons
    nombre_total_taches_non_recap_jalon = len(taches_non_recap_jalon)

    # Critères existants :
    recap_avec_both = 0  # Tâches récapitulatives avec à la fois des prédécesseurs et des successeurs
    recap_avec_successeur = 0  # Tâches récapitulatives avec des successeurs
    recap_avec_predecesseur = 0  # Tâches récapitulatives avec des prédécesseurs
    recap_avec_autre_contrainte = 0  # Tâches récapitulatives avec contrainte autre que "Dès Que Possible"
    recap_et_jalon = 0  # Critère pour "Récapitulatif et Jalon"
    non_recap_jalon_date_fixee = 0  # Critère pour "Date fixée non récapitulatif et jalon"
    non_recap_non_jalon_date_fixee = 0  # Critère pour "Date fixée non récapitulatif et non jalon"
    non_recap_duree_sup_60 = 0 
    non_recap_predecesseur_retard_avance = 0  # Critère pour "Tâches non-récap avec prédécesseur en avance ou en retard"
    non_recap_predecesseur_avance = 0  # Critère pour avance
    non_recap_predecesseur_retard = 0  # Critère pour retard
    non_recap_marge_totale_sup_60 = 0 # Critète marge totale
    non_recap_non_jalon_absence_pred_succ = 0  # Tâches non récapitulatives, non jalons, et sans prédécesseur ni successeur
    non_recap_non_jalon_absence_predecesseur = 0  # Tâches non récapitulatives, non jalons, et sans prédécesseur
    non_recap_non_jalon_absence_successeur = 0  # Tâches non récapitulatives, non jalons, et sans successeur
    date_fixee_occurrences = 0  # Toutes les tâches dont la contrainte est autre que "Dès Que Possible" ou "Le Plus Tard Possible"
    recap_date_fixee = 0  # Tâches récapitulatives avec "Date fixée"
    non_recap_date_fixee = 0  # Tâches non récapitulatives avec "Date fixée"

    # Parcourir toutes les tâches
    for id_tache, wbs, successeurs, predecesseurs, type_contrainte, duree, date_debut, date_fin, marge_totale in taches:
        # Critères pour tâches récapitulatives
        if id_tache in taches_recap:
            if successeurs or predecesseurs:
                recap_avec_both += 1
            if successeurs:
                recap_avec_successeur += 1
            if predecesseurs:
                recap_avec_predecesseur += 1
            if type_contrainte != "Dès Que Possible":
                recap_avec_autre_contrainte += 1
            if type_contrainte not in ["Dès Que Possible", "Le Plus Tard Possible"]:
                recap_date_fixee += 1
            if isinstance(duree, str) and duree.startswith("0"):
                recap_et_jalon += 1
        
        # Vérifier si la tâche est un jalon
        est_jalon = isinstance(duree, str) and duree.startswith("0")  # Vérifie si la durée commence par "0" (0j, 0jr, 0jour, etc.)

        # Critères pour tâches non récapitulatives et non jalons
        if id_tache in taches_non_recap and not est_jalon:
            if not successeurs or not predecesseurs:
                non_recap_non_jalon_absence_pred_succ += 1
            if not predecesseurs:
                non_recap_non_jalon_absence_predecesseur += 1
            if not successeurs:
                non_recap_non_jalon_absence_successeur += 1
            if type_contrainte not in ["Dès Que Possible", "Le Plus Tard Possible"]:
                non_recap_date_fixee += 1

        # Critère "Date fixée" pour les tâches non récapitulatives et jalons
        if id_tache in taches_non_recap and est_jalon and type_contrainte not in ["Dès Que Possible", "Le Plus Tard Possible"]:
            non_recap_jalon_date_fixee += 1

        # Critère "Date fixée" pour les tâches non récapitulatives et non-jalons
        if id_tache in taches_non_recap and not est_jalon and type_contrainte not in ["Dès Que Possible", "Le Plus Tard Possible"]:
            non_recap_non_jalon_date_fixee += 1

        # Critère "Date fixée" : Si la contrainte n'est ni "Dès Que Possible" ni "Le Plus Tard Possible"
        if type_contrainte not in ["Dès Que Possible", "Le Plus Tard Possible"]:
            date_fixee_occurrences += 1

        # Critère "Durée supérieure à 60 jours ouvrés"
        duree_jours_ouvres = calculer_jours_ouvres(date_debut, date_fin)
        if duree_jours_ouvres > 60:
            non_recap_duree_sup_60 += 1

        # Critères pour tâches non récapitulatives
        if id_tache in taches_non_recap:
            avance_count = 0
            retard_count = 0
            marge_numerique = int(''.join(filter(str.isdigit, marge_totale)))
            taches_marge_negative = 0 

            if predecesseurs:
                # Compter le nombre de "+" et "-"
                avance_count = predecesseurs.count("+")
                retard_count = predecesseurs.count("-")
                non_recap_predecesseur_retard_avance += avance_count + retard_count
            
            # Critères séparés pour l'avance et le retard
            non_recap_predecesseur_avance += avance_count
            non_recap_predecesseur_retard += retard_count
        # Critère : Si la marge totale dépasse 60 jours
            if marge_numerique > 60:
                non_recap_marge_totale_sup_60 += 1

            if "-" in marge_totale:
                taches_marge_negative += 1


    #Calcul du nombre total pour ratio 
    nombre_total_taches_recap = len(taches_recap) # Nombre total de tâches récapitulatives
    nombre_total_taches_non_recap = len(taches_non_recap)
    nombre_total_taches_non_recap_non_jalon =  len(taches_non_recap_non_jalon)
    nombre_total_taches_non_recap_jalon = len(taches_non_recap_jalon)
    nombre_total_taches = len(taches)  # Nombre total de tâches, récapitulatives ou non

    # Calcul des ratios pour les critères récapitulatives
    ratio_avec_both = (recap_avec_both / nombre_total_taches_recap) * 100 if nombre_total_taches_recap > 0 else 0
    ratio_avec_successeur = (recap_avec_successeur / nombre_total_taches_recap) * 100 if nombre_total_taches_recap > 0 else 0
    ratio_avec_predecesseur = (recap_avec_predecesseur / nombre_total_taches_recap) * 100 if nombre_total_taches_recap > 0 else 0
    ratio_autre_contrainte = (recap_avec_autre_contrainte / nombre_total_taches_recap) * 100 if nombre_total_taches_recap > 0 else 0
    ratio_recap_date_fixee = (recap_date_fixee / nombre_total_taches_recap) * 100 if nombre_total_taches_recap > 0 else 0
    ratio_recap_et_jalon = (recap_et_jalon / nombre_total_taches_recap) * 100 if nombre_total_taches_recap > 0 else 0

    # Calcul des ratios pour les nouveaux critères non récapitulatives
    ratio_non_recap_pred_succ = (non_recap_non_jalon_absence_pred_succ / nombre_total_taches_non_recap_non_jalon) * 100 if nombre_total_taches_non_recap_non_jalon > 0 else 0
    ratio_non_recap_predecesseur = (non_recap_non_jalon_absence_predecesseur / nombre_total_taches_non_recap_non_jalon) * 100 if nombre_total_taches_non_recap_non_jalon > 0 else 0
    ratio_non_recap_successeur = (non_recap_non_jalon_absence_successeur / nombre_total_taches_non_recap_non_jalon) * 100 if nombre_total_taches_non_recap_non_jalon > 0 else 0
    ratio_non_recap_date_fixee = (non_recap_date_fixee / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0
    ratio_non_recap_jalon_date_fixee = (non_recap_jalon_date_fixee / nombre_total_taches_non_recap_jalon) * 100 if nombre_total_taches_non_recap_jalon > 0 else 0
    ratio_non_recap_non_jalon_date_fixee = (non_recap_non_jalon_date_fixee / nombre_total_taches_non_recap_non_jalon) * 100 if nombre_total_taches_non_recap_non_jalon > 0 else 0
    ratio_non_recap_duree_sup_60 = (non_recap_duree_sup_60 / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0
    ratio_non_recap_predecesseur_retard_avance = (non_recap_predecesseur_retard_avance / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0
    ratio_non_recap_predecesseur_avance = (non_recap_predecesseur_avance / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0
    ratio_non_recap_predecesseur_retard = (non_recap_predecesseur_retard / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0
    
    # Ratio pour "Date fixée" (global)
    ratio_date_fixee = (date_fixee_occurrences / nombre_total_taches) * 100 if nombre_total_taches > 0 else 0

    # Calcul du ratio pour le critère "marge totale supérieure à 60 jours"
    ratio_marge_totale_sup_60 = (non_recap_marge_totale_sup_60 / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0

    # Calcul du ratio pour les tâches avec marge négative
    ratio_marge_negative = (taches_marge_negative / nombre_total_taches_non_recap) * 100 if nombre_total_taches_non_recap > 0 else 0


    # Seuils et indicateurs 
    seuil_avec_both = "0% = 5 / <=2% = 3 / >2% = 1"
    indicateur_avec_both = 5 if ratio_avec_both == 0 else (3 if ratio_avec_both <= 2 else 1)
    
    seuil_avec_successeur = "0% = 5 / <=2% = 3 / >2% = 1"
    indicateur_avec_successeur = 5 if ratio_avec_successeur == 0 else (3 if ratio_avec_successeur <= 2 else 1)
    
    seuil_avec_predecesseur = "0% = 5 / <=2% = 3 / >2% = 1"
    indicateur_avec_predecesseur = 5 if ratio_avec_predecesseur == 0 else (3 if ratio_avec_predecesseur <= 2 else 1)
    
    seuil_autre_contrainte = "0% = 5 / <=2% = 3 / >2% = 1"
    indicateur_autre_contrainte = 5 if ratio_autre_contrainte == 0 else (3 if ratio_autre_contrainte <= 2 else 1)

    seuil_recap_jalon = "0% ou 1"
    indicateur_recap_jalon = 5 if ratio_recap_et_jalon == 0 else 1

    seuil_non_recap_pred_succ = "0% = 5 / <=2% = 3 / >2% = 1"
    indicateur_non_recap_pred_succ = 5 if ratio_non_recap_pred_succ == 0 else (3 if ratio_non_recap_pred_succ <= 2 else 1)

    seuil_non_recap_predecesseur = "5 <2% / 3 <=5% / 1 >5%"
    indicateur_non_recap_predecesseur = 5 if ratio_non_recap_predecesseur <= 2 else (3 if ratio_non_recap_predecesseur <= 5 else 1)

    seuil_non_recap_successeur = "5 <2% / 3 <=5% / 1 >5%"
    indicateur_non_recap_successeur = 5 if ratio_non_recap_successeur <= 2 else (3 if ratio_non_recap_successeur <= 5 else 1)

    seuil_date_fixee = "N/A"
    
    seuil_recap_date_fixee = "0% = 5 / <=2% = 3 / >2% = 1"
    indicateur_recap_date_fixee = 5 if ratio_recap_date_fixee == 0 else (3 if ratio_recap_date_fixee <= 2 else 1)

    seuil_non_recap_date_fixee = "N/A"
    #indicateur_non_recap_date_fixee = 5 if ratio_non_recap_date_fixee == 0 else (3 if ratio_non_recap_date_fixee <= 2 else 1)

    seuil_non_recap_jalon_date_fixee = "N/A"

    seuil_non_recap_non_jalon_date_fixee = "5 <= 5% / 3<=10% / 1 >10%"
    indicateur_non_recap_non_jalon_date_fixee = 5 if ratio_non_recap_non_jalon_date_fixee <= 5 else (3 if ratio_non_recap_non_jalon_date_fixee <= 10 else 1)

    seuil_duree_sup_60 = "5 <= 5% / 3<=10% / 1 >10%"
    indicateur_duree_sup_60 = 5 if ratio_non_recap_duree_sup_60 <= 5 else (3 if ratio_non_recap_duree_sup_60 <= 10 else 1)

    seuil_non_recap_predecesseur_retard_avance = "5 <3% / 3 <=5% / 1 >5%"
    indicateur_non_recap_predecesseur_retard_avance = 5 if ratio_non_recap_predecesseur_retard_avance < 3 else (3 if ratio_non_recap_predecesseur_retard_avance <= 5  else 1)

    seuil_non_recap_predecesseur_avance = "5 = 1% / 3 <=2% / 1 >2%"
    indicateur_non_recap_predecesseur_avance = 5 if ratio_non_recap_predecesseur_avance == 1 else (3 if ratio_non_recap_predecesseur_avance <= 2 else 1)

    seuil_non_recap_predecesseur_retard = "5 <3% / 3 <=8% / 1 >8%"
    indicateur_non_recap_predecesseur_retard = 5 if ratio_non_recap_predecesseur_retard < 3 else (3 if ratio_non_recap_predecesseur_retard <= 8 else 1)

    seuil_marge_totale_sup_60 = "5 <2% / 3 <=5% / 1 >5%"
    indicateur_marge_totale_sup_60 = 5 if ratio_marge_totale_sup_60 < 2 else (3 if ratio_marge_totale_sup_60 <= 5 else 1)

    seuil_marge_negative = "5 = 0% / 3 <=2% / 1 >2%"
    indicateur_marge_negative = 5 if ratio_marge_negative == 0 else (3 if ratio_marge_negative <= 2 else 1)

    # Résultats des critères
    resultats = [
        {
            "critere": "Tâches récapitulatives avec prédécesseurs et successeurs",
            "nombre_occurences": recap_avec_both,
            "nombre_total": nombre_total_taches_recap,
            "ratio": ratio_avec_both,
            "seuil": seuil_avec_both,
            "indicateur": indicateur_avec_both,
        },
        {
            "critere": "    Tâches récapitulatives avec successeurs",
            "nombre_occurences": recap_avec_successeur,
            "nombre_total": nombre_total_taches_recap,
            "ratio": ratio_avec_successeur,
            "seuil": seuil_avec_successeur,
            "indicateur": indicateur_avec_successeur,
        },
        {
            "critere": "    Tâches récapitulatives avec prédécesseurs",
            "nombre_occurences": recap_avec_predecesseur,
            "nombre_total": nombre_total_taches_recap,
            "ratio": ratio_avec_predecesseur,
            "seuil": seuil_avec_predecesseur,
            "indicateur": indicateur_avec_predecesseur,
        },
        {
            "critere": "Tâches récapitulatives avec contrainte autre que 'Dès Que Possible'",
            "nombre_occurences": recap_avec_autre_contrainte,
            "nombre_total": nombre_total_taches_recap,
            "ratio": ratio_autre_contrainte,
            "seuil": seuil_autre_contrainte,
            "indicateur": indicateur_autre_contrainte,
        },
        {
            "critere": "    Non-récapitulatif et Non Jalon + absence de prédécesseur ou successeur",
            "nombre_occurences": non_recap_non_jalon_absence_pred_succ,
            "nombre_total": nombre_total_taches_non_recap_non_jalon,
            "ratio": ratio_non_recap_pred_succ,
            "seuil": seuil_non_recap_pred_succ,
            "indicateur": indicateur_non_recap_pred_succ,
        },
        {
            "critere": "    Non-récapitulatif et Non Jalon + absence de prédécesseur",
            "nombre_occurences": non_recap_non_jalon_absence_predecesseur,
            "nombre_total": nombre_total_taches_non_recap_non_jalon,
            "ratio": ratio_non_recap_predecesseur,
            "seuil": seuil_non_recap_predecesseur,
            "indicateur": indicateur_non_recap_predecesseur,
        },
        {
            "critere": "    Non-récapitulatif et Non Jalon + absence de successeur",
            "nombre_occurences": non_recap_non_jalon_absence_successeur,
            "nombre_total": nombre_total_taches_non_recap_non_jalon,
            "ratio": ratio_non_recap_successeur,
            "seuil": seuil_non_recap_successeur,
            "indicateur": indicateur_non_recap_successeur,
        },
        {
            "critere": "Recapitulatif + Jalon",
            "nombre_occurences": recap_et_jalon,
            "nombre_total": nombre_total_taches_recap,
            "ratio": ratio_recap_et_jalon,
            "seuil": seuil_recap_jalon,
            "indicateur": indicateur_recap_jalon,
        },
        {
            "critere": "Tâches avec 'Date fixée' ",
            "nombre_occurences": date_fixee_occurrences,
            "nombre_total": nombre_total_taches,
            "ratio": ratio_date_fixee,
            "seuil": seuil_date_fixee,
            "indicateur": "",
        },
        {
            "critere": "    Tâches récapitulatives avec 'Date fixée' (ni 'Dès Que Possible' ni 'Le Plus Tard Possible')",
            "nombre_occurences": recap_date_fixee,
            "nombre_total": nombre_total_taches_recap,
            "ratio": ratio_recap_date_fixee,
            "seuil": seuil_recap_date_fixee,
            "indicateur": indicateur_recap_date_fixee,
        },
        {
            "critere": "    Tâches non-récapitulatives avec 'Date fixée' (ni 'Dès Que Possible' ni 'Le Plus Tard Possible')",
            "nombre_occurences": non_recap_date_fixee,
            "nombre_total": nombre_total_taches_non_recap,
            "ratio": ratio_non_recap_date_fixee,
            "seuil": seuil_non_recap_date_fixee,
            "indicateur": "",
        },
        {
            "critere": "        Date fixée non récapitulatif et jalon",
            "nombre_occurences": non_recap_jalon_date_fixee,
            "nombre_total": nombre_total_taches_non_recap_jalon,
            "ratio": ratio_non_recap_jalon_date_fixee,
            "seuil": seuil_non_recap_jalon_date_fixee,
            "indicateur": "",
        },
        {
            "critere": "        Date fixée non récapitulatif et non jalon",
            "nombre_occurences": non_recap_non_jalon_date_fixee,
            "nombre_total": nombre_total_taches_non_recap_non_jalon,
            "ratio": ratio_non_recap_non_jalon_date_fixee,
            "seuil": seuil_non_recap_non_jalon_date_fixee,
            "indicateur": indicateur_non_recap_non_jalon_date_fixee,
        },
        {
            "critere": "Non-récapitulatif avec Durée supérieure à 60 jours",
            "nombre_occurences": non_recap_duree_sup_60,
            "nombre_total": nombre_total_taches_non_recap,
            "ratio": ratio_non_recap_duree_sup_60,
            "seuil": seuil_duree_sup_60,
            "indicateur": indicateur_duree_sup_60,
        },

        {
            "critere": "Tâches non-récap avec prédécesseur en avance ou en retard",
            "nombre_occurences": non_recap_predecesseur_retard_avance,
            "nombre_total": nombre_total_taches_non_recap,
            "ratio": ratio_non_recap_predecesseur_retard_avance,
            "seuil": seuil_non_recap_predecesseur_retard_avance,
            "indicateur": indicateur_non_recap_predecesseur_retard_avance,
        },
        {
            "critere": "    Tâches non-récap avec prédécesseur en avance",
            "nombre_occurences": non_recap_predecesseur_avance,
            "nombre_total": nombre_total_taches_non_recap,
            "ratio": ratio_non_recap_predecesseur_avance,
            "seuil": seuil_non_recap_predecesseur_avance,
            "indicateur": indicateur_non_recap_predecesseur_avance,
        },
        {
            "critere": "    Tâches non-récap avec prédécesseur en retard",
            "nombre_occurences": non_recap_predecesseur_retard,
            "nombre_total": nombre_total_taches_non_recap,
            "ratio": ratio_non_recap_predecesseur_retard,
            "seuil": seuil_non_recap_predecesseur_retard,
            "indicateur": indicateur_non_recap_predecesseur_retard,
        },

        {
            "critere": "Tâches non récapitulatives avec marge totale > 60 jours",
            "nombre_occurences": non_recap_marge_totale_sup_60,
            "nombre_total": nombre_total_taches_non_recap,
            "ratio": ratio_marge_totale_sup_60,
            "seuil": seuil_marge_totale_sup_60,
            "indicateur": indicateur_marge_totale_sup_60,
        },
        {
        "critere": "Tâches avec marge négative",
        "nombre_occurences": taches_marge_negative,
        "nombre_total": nombre_total_taches_non_recap,
        "ratio": ratio_marge_negative,
        "seuil": seuil_marge_negative,
        "indicateur": indicateur_marge_negative,
        }
    ]

    return resultats





# Fonction pour évaluer le projet sélectionné et afficher le rapport
def afficher_evaluation(projets_dict, combo_projets, tree):
    projet_selectionne = combo_projets.get()
    if not projet_selectionne:
        messagebox.showwarning("Erreur", "Veuillez sélectionner un projet.")
        return
    
    projet_id = projets_dict[projet_selectionne]
    
    # Évaluer les critères du projet
    resultats = evaluer_projet(projet_id)

    # Effacer les anciens résultats
    tree.delete(*tree.get_children())

    # Ajouter les résultats au tableau avec la coloration en fonction de l'indicateur
    for resultat in resultats:
        # Appliquer un tag en fonction de l'indicateur
        if resultat["indicateur"] == 1:
            tag = 'rouge'
        elif resultat["indicateur"] == 3:
            tag = 'orange'
        elif resultat["indicateur"] == 5:
            tag = 'vert'
        else:
            tag = ''  # Aucun tag si l'indicateur ne correspond pas à un cas précis

        # Insérer la ligne dans le Treeview avec le tag approprié
        tree.insert("", "end", values=(
            resultat["critere"],
            resultat["nombre_occurences"],
            resultat["nombre_total"],
            round(resultat["ratio"], 2),
            resultat["seuil"],
            resultat["indicateur"]
        ), tags=(tag,))


#--------------------------------------------------------------
# Création de la fenêtre principale
root = tk.Tk()
root.title("Importation de Planning")

# Mettre la fenêtre en mode "agrandi" (comme si l'utilisateur avait cliqué sur le bouton d'agrandissement)
root.state('zoomed')

# Style ttk pour rendre les widgets plus modernes
style = ttk.Style()
style.theme_use('clam')  # Thème moderne

# Charger l'image du logo 
image = Image.open("LGM_logo.png")
image = image.resize((150, 100), Image.LANCZOS)  # Redimensionner l'image pour qu'elle prenne moins de place
logo_image = ImageTk.PhotoImage(image)

# Personnaliser les boutons et autres widgets
style.configure('TButton', background='#007ACC', foreground='white', font=('Arial', 12, 'bold'))
style.configure('TLabel', font=('Arial', 12), padding=10)
style.configure('TEntry', padding=5)

# Création du frame principal qui contient tout
main_frame = ttk.Frame(root, padding="10 10 10 10")
main_frame.pack(expand=True, fill='both')

# Frame pour le logo et le titre
logo_title_frame = ttk.Frame(main_frame)
logo_title_frame.pack(side='top', fill='x', pady=10)

# Ajouter le logo en haut à gauche (première colonne de la grille)
logo_label = ttk.Label(logo_title_frame, image=logo_image)
logo_label.grid(row=0, column=0, padx=20, pady=10, sticky='w')  # Placer à gauche avec sticky et ajouter du padding

# Ajouter un titre à côté du logo (se positionnera au centre)
titre_label = ttk.Label(logo_title_frame, text="Évaluateur de planning", font=("Arial", 45, "bold"))
titre_label.grid(row=0, column=1, padx=(10, 10), pady=(30, 0), sticky='nsew')  # Centré avec padding ajusté

# Ajuster les proportions des colonnes pour centrer correctement le logo et le titre
logo_title_frame.grid_columnconfigure(0, weight=1)  # Première colonne pour le logo
logo_title_frame.grid_columnconfigure(1, weight=2)  # Deuxième colonne pour le titre

# Frame pour contenir et centrer le bloc de widgets
bloc_widgets_frame = ttk.Frame(main_frame)
bloc_widgets_frame.pack(expand=True, pady=50)  # Ajouter de l'espace autour, 50 pixels en haut et en bas

# Ajout du titre pour le nom du projet
label_nom_projet = ttk.Label(bloc_widgets_frame, text="Nom du projet :", style='TLabel')
label_nom_projet.pack(pady=(20, 5), anchor='center')  # Ajuster le padding pour un bon espacement

# Champ d'entrée pour le nom du projet (augmenter la hauteur et la taille de la police)
entry_nom_projet = ttk.Entry(bloc_widgets_frame, text="Nom du projet :", width=50, font=("Arial", 14))  # Police plus grande
entry_nom_projet.pack(pady=(10, 20), anchor='center')  # Ajuster le padding entre le label et l'entry

# Bouton pour importer le fichier (taille augmentée)
btn_importer = ttk.Button(bloc_widgets_frame, text="Importer un fichier Excel", command=importer_fichier, width=35)
btn_importer.pack(pady=15, anchor='center')

# Bouton pour traiter le fichier (taille augmentée)
btn_traiter = ttk.Button(bloc_widgets_frame, text="Traiter et envoyer le fichier", command=traiter_fichier, state="disabled", width=35)
btn_traiter.pack(pady=15, anchor='center')

# Barre de progression (agrandir et ajuster la couleur)
progress_bar = ttk.Progressbar(bloc_widgets_frame, orient="horizontal", length=600, mode="determinate")
progress_bar.pack(pady=30, anchor='center')  # Plus d'espace autour de la barre de progression

# Cadre pour le bouton en bas (toujours dans le main_frame)
bottom_frame = ttk.Frame(main_frame)
bottom_frame.pack(side='bottom', fill='x', padx=30, pady=20)  # Réduire l'espace en bas et à droite

# Bouton pour accéder à la page des rapports
btn_aller_rapport = ttk.Button(bottom_frame, text="Aller aux rapports", command=afficher_page_evaluation, width=20)
btn_aller_rapport.pack(side='right', padx=10, pady=10)

# Centrer la fenêtre
def centrer_fenetre(root):
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

centrer_fenetre(root)

# Assurer que les widgets soient redimensionnés proprement lors du redimensionnement de la fenêtre
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Lancer l'interface
root.mainloop()

# Fermer la connexion à la base de données après la fermeture de l'interface
conn.close()
