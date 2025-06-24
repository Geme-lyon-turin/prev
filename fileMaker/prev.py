import os

import pandas as pd
from datetime import datetime
import unicodedata
import re
import tkinter as tk
import sys

class RedirectText:
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, message):
        self.output.insert(tk.END, message)
        self.output.see(tk.END)  # scroll automatique

    def flush(self):  # Nécessaire pour être compatible avec sys.stdout
        pass

def CO9():
    today = datetime.today().strftime("%d/%m/%Y")

    # 🧹 Fonction de nettoyage brut
    def nettoyer_texte(texte):
        if pd.isna(texte):
            return ""
        return str(texte).replace('\n', ' ').replace('\r', '').strip()

    # 🔧 Fonction de normalisation robuste (minuscules, accents, espaces)
    def normaliser_texte(txt):
        txt = nettoyer_texte(txt)
        txt = unicodedata.normalize('NFD', txt).encode('ascii', 'ignore').decode("utf-8")
        txt = txt.lower()
        txt = re.sub(r"\s+", " ", txt)  # supprime les espaces multiples
        return txt.strip()

    # 📂 Lecture du fichier Excel source
    for nom_fichier in os.listdir("./data"):
        if "MT103" in nom_fichier and nom_fichier:
            file_path = os.path.join("./data", nom_fichier)
    sheet_name = "Planning besoins t"

    df = pd.read_excel(file_path, header=13, sheet_name=sheet_name)
    df.columns = [nettoyer_texte(col) for col in df.columns]  # Nettoyage des noms de colonnes

    # 📌 Lecture du fichier de correspondance des colonnes à extraire
    colonnes_valeurs = pd.read_csv("./mapping/correspondance_CO9_materiaux.csv", header=None)
    colonnes_valeurs = colonnes_valeurs.iloc[:, 0].apply(nettoyer_texte).tolist()

    # Normalisation pour matcher plus facilement
    colonnes_df_normalisees = {normaliser_texte(col): col for col in df.columns}
    colonnes_mapping_normalisees = [normaliser_texte(col) for col in colonnes_valeurs]

    # Liste finale des colonnes existantes à traiter
    colonnes_a_traiter = []
    for col_norm in colonnes_mapping_normalisees:
        if col_norm in colonnes_df_normalisees:
            colonnes_a_traiter.append(colonnes_df_normalisees[col_norm])

    # 📥 Chargement du mapping des champs de sortie
    mapping_df = pd.read_csv("mapping/CO9_colonnes.csv", dtype=str)
    mapping_dict = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))

    # 🧪 Transformation des données
    donnees_transformees = []
    # 📥 Chargement de la correspondance Typo courte
    correspondance_df = pd.read_csv("./mapping/correspondance_CO9_materiaux.csv")
    correspondance_df["Typo extraction normalisee"] = correspondance_df["Typo extraction"].apply(normaliser_texte)
    # Création d'un dictionnaire {colonne_normalisee: valeur_typo_courte}
    typo_courte_dict = dict(zip(correspondance_df["Typo extraction normalisee"], correspondance_df["Typo courte"]))

    for index, row in df.iterrows():
        for col in colonnes_a_traiter:
            valeur_brute = row[col]
            if col not in df.columns:
                print(f"⚠️ Colonne absente : {col}")
                continue
            try:
                valeur_num = float(str(valeur_brute).replace(",", "."))
                if valeur_num > 0:
                    nouvelle_ligne = {}

                    for champ_cible, source in mapping_dict.items():
                        source_str = str(source).strip()

                        if source_str == "valeur":
                            nouvelle_ligne[champ_cible] = round(valeur_num, 2)
                        elif source_str == "":
                            nouvelle_ligne[champ_cible] = ""
                        elif source_str == "today":
                            nouvelle_ligne[champ_cible] = today
                        elif source_str == "intitulé_rapport":
                            annee = ""
                            semaine = ""

                            # Récupération année
                            if "Année" in row and pd.notna(row["Année"]):
                                try:
                                    annee = str(int(float(row["Année"])))
                                except:
                                    annee = ""

                            # Récupération semaine +1
                            if "Semaine" in row and pd.notna(row["Semaine"]):
                                try:
                                    semaine_num = int(float(row["Semaine"])) + 1
                                    semaine = f"S{semaine_num:02d}"
                                except:
                                    semaine = ""

                            nouvelle_ligne[champ_cible] = f"CO9 {annee} {semaine}".strip()

                        elif source_str in ["head", "colonne_name"]:
                            nouvelle_ligne[champ_cible] = col
                        elif source_str == "Typo courte":
                            col_normalise = normaliser_texte(col)
                            nouvelle_ligne[champ_cible] = typo_courte_dict.get(col_normalise, "")
                        elif source_str == "Année":
                            nouvelle_ligne[champ_cible] = str(int(row["Année"])) if "Année" in row and pd.notna(
                                row["Année"]) else ""
                        elif source_str == "Mois":
                            nouvelle_ligne[champ_cible] = str(int(row["Mois"])) if "Mois" in row and pd.notna(
                                row["Mois"]) else ""
                        elif source_str == "Semaine":
                            nouvelle_ligne[champ_cible] = str(int(row["Semaine"]+1)) if "Semaine" in row and pd.notna(
                                row["Semaine"]) else ""
                        elif source_str == "SemaineS":
                            nouvelle_ligne[champ_cible] = f"S{int(row['Semaine']+1):02}" if "Semaine" in row and pd.notna(row["Semaine"]) else ""
                        elif source_str == "Niveau précision":
                            nouvelle_ligne[champ_cible] = str(
                                row["Niveau précision"]).strip() if "Niveau précision" in row and pd.notna(row["Niveau précision"]) else ""
                        elif source_str.isdigit():
                            try:
                                nouvelle_ligne[champ_cible] = str(int(row.iloc[int(source_str)]))
                            except Exception:
                                nouvelle_ligne[champ_cible] = ""
                        else:
                            nouvelle_ligne[champ_cible] = source_str

                    donnees_transformees.append(nouvelle_ligne)

            except Exception as e:
                print(f"❌ Impossible de convertir la valeur : {valeur_brute} → {e}")
                continue

    # 📤 Export CSV
    df_final = pd.DataFrame(donnees_transformees)

    output_file = f"CO9_BDD_UNIFIEE_{datetime.today().date()}.xlsx"
    df_final.to_excel(output_file, index=False)

    print(f"\n✅ Fichier BDD_UNIFIEE_CO9 généré ✅ :")

def CO8():
    # 📂 Charger le fichier brut pour accéder à la ligne 6 (index 5)
    for nom_fichier in os.listdir("./data"):
        if "CO8" in nom_fichier and nom_fichier:
            file_path = os.path.join("./data", nom_fichier)
    df_brut = pd.read_excel(file_path, header=None)

    # 🧠 Extraire la ligne 6 (Typo_Flux) AVANT le vrai header
    typo_flux_ligne = df_brut.iloc[4]

    # 📂 Charger le vrai DataFrame avec header à la ligne 7 (index 6)
    df = pd.read_excel(file_path, header=6)

    # 🔁 Charger le mapping des colonnes
    mapping_df = pd.read_csv("mapping/CO8_colonnes.csv", dtype=str)
    mapping_dict = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))

    # ✅ Colonnes avec des valeurs numériques à traiter
    colonnes_valeurs = ['0/4', '4/8', '8/16', '0/4.1', '0/4 CL', '4/10', '10/20', 'eFs', 'eFsg', 'l1-42', 't7SB3',
                        'CL1', 'CL2', 'CL3a', 'CL3b', 'CL1.1', 'CL2.1', 'CL3a.1', 'CL3b.1']

    # 🛠 Préparation pour la transformation
    donnees_transformees = []
    today = datetime.today().date().strftime("%d/%m/%Y")
    valeur_head_precedente = None

    # 🔄 Traitement des lignes
    for _, row in df.iterrows():
        for col in colonnes_valeurs:
            valeur = row[col]
            if pd.notna(valeur) and isinstance(valeur, (int, float)):
                nouvelle_ligne = {}
                numero_colonne = df.columns.get_loc(col)
                for champ_cible, source in mapping_dict.items():
                    source_str = str(source).strip()

                    # Normalisation du nom de colonne
                    col_base = "0/4" if col == "0/4.1" else col

                    if source_str == "valeur":
                        nouvelle_ligne[champ_cible] = valeur
                    elif source_str == "":
                        nouvelle_ligne[champ_cible] = ""
                    elif source_str == "intitulé_rapport":
                        annee = ""
                        try:
                            val_annee = row.iloc[1]
                            if pd.notna(val_annee):
                                annee = str(int(val_annee))
                        except:
                            annee = ""

                        semaine = ""
                        try:
                            val_semaine = row.iloc[3]
                            if pd.notna(val_semaine):
                                semaine_num = int(val_semaine)
                                semaine = f"S{semaine_num:02d}"
                        except:
                            semaine = ""

                        nouvelle_ligne[champ_cible] = f"CO8 {annee} {semaine}".strip()

                    elif source_str == "today":
                        nouvelle_ligne[champ_cible] = today
                    elif source_str == "head":
                        nouvelle_ligne[champ_cible] = col_base
                    elif source_str == "s3":
                        try:
                            val = row.iloc[3]
                            if pd.notna(val):
                                semaine_num = int(val)
                                nouvelle_ligne[champ_cible] = f"S{semaine_num:02}"  # ex: 6 → S06
                            else:
                                nouvelle_ligne[champ_cible] = ""
                        except Exception:
                            nouvelle_ligne[champ_cible] = ""
                    elif source_str == "1" or source_str == "3":
                        val = row.iloc[int(source_str)]
                        try:
                            nouvelle_ligne[champ_cible] = str(int(val))  # Convertit 2025.0 → "2025"
                        except Exception:
                            nouvelle_ligne[champ_cible] = ""
                    elif source_str == "4":
                        val = row.iloc[int(source_str)]
                        try:
                            if isinstance(val, (pd.Timestamp, datetime)):
                                nouvelle_ligne[champ_cible] = val.strftime("%d")  # ex: "04"
                            else:
                                # Tente de parser la valeur si ce n'est pas un datetime
                                date_obj = pd.to_datetime(val)
                                nouvelle_ligne[champ_cible] = date_obj.strftime("%d")
                        except Exception:
                            nouvelle_ligne[champ_cible] = ""
                    elif source_str == "head-1":
                        i = 0
                        valeur_typo = ""
                        while (numero_colonne - i) >= 0:
                            candidat = typo_flux_ligne[numero_colonne - i]
                            if pd.notna(candidat) and str(candidat).strip() != "":
                                valeur_typo = candidat
                                break
                            i += 1
                        nouvelle_ligne[champ_cible] = valeur_typo
                    elif source_str.isdigit():
                        nouvelle_ligne[champ_cible] = row.iloc[int(source_str)]
                    else:
                        nouvelle_ligne[champ_cible] = source_str
                donnees_transformees.append(nouvelle_ligne)

    # ✅ Création du DataFrame final
    df_final = pd.DataFrame(donnees_transformees)

    # 💾 Sauvegarde en CSV
    df_final.to_excel(f"CO8_BDD_UNIFIEE_{datetime.today().date()}.xlsx", index=False)

    print("✅ Fichier BDD_UNIFIEE_CO8 généré ✅")

def CO67():
    def get_raw_header(chantier):
        for nom_fichier in os.listdir("./data"):
            if "CO67" in nom_fichier and nom_fichier:
                file_path = os.path.join("./data", nom_fichier)
        raw_header = pd.read_excel(file_path, header=None, sheet_name=chantier)
        # lignes importantes
        ligne_entete = raw_header.iloc[5]  # index 5 = ligne 6 Excel
        pk = raw_header.iloc[4]  # index 3 = ligne 4 Excel
        site = raw_header.iloc[1]  # index 0 = ligne 1 Excel
        ouvrage = raw_header.iloc[2]  # index 1 = ligne 2 Excel
        return ligne_entete, pk, site, ouvrage

    def get_CO6():
        for nom_fichier in os.listdir("./data"):
            if "CO67" in nom_fichier and nom_fichier:
                file_path = os.path.join("./data", nom_fichier)
        df = pd.read_excel(file_path, header=6, sheet_name="CO6")
        df['Jour'] = df["Pré classement\nCorrigé à front"].astype(str).str.strip()
        df['Date'] = pd.to_datetime(df["Unnamed: 1"], errors='coerce')
        df = df[df['Date'].notna()]
        return df

    def get_CO7():
        for nom_fichier in os.listdir("./data"):
            if "CO67" in nom_fichier and nom_fichier:
                file_path = os.path.join("./data", nom_fichier)
        df = pd.read_excel(file_path, header=6, sheet_name="CO7")
        df['Jour'] = df["Pré classement\nCorrigé à front"].astype(str).str.strip()
        df['Date'] = pd.to_datetime(df["Unnamed: 1"], errors='coerce')
        df = df[df['Date'].notna()]
        return df

    def formatage_data(df, chantier):
        colonnes_valeurs = ['cl1', 'cl2', 'cl3']  # à adapter selon ton cas
        mapping_df = pd.read_csv("mapping/CO6-7_colonnes.csv")
        mapping_dict = dict(zip(mapping_df["Source"], mapping_df["Destination"]))

        ligne_entete, pk, site, ouvrage = get_raw_header(chantier)

        today = datetime.today().strftime("%d/%m/%Y")
        donnees_transformees = []

        for i, row in df.iterrows():
            date_obj = row['Date']
            colonnes_utiles = [col for col in df.columns if any(val.lower() in col.lower() for val in colonnes_valeurs)]

            for col in colonnes_utiles:
                valeur = row[col]
                if pd.notna(valeur) and isinstance(valeur, (int, float)):
                    nouvelle_ligne = {}

                    for champ_cible, source in mapping_dict.items():
                        if source == "valeur":
                            nouvelle_ligne[champ_cible] = round(valeur, 2)
                        elif source == "Chantier":
                            nouvelle_ligne[champ_cible] = chantier
                        elif source == "":
                            nouvelle_ligne[champ_cible] = ""
                        elif source == "today":
                            nouvelle_ligne[champ_cible] = today
                        elif source == "intitulé_rapport":
                            try:
                                annee = date_obj.year
                                semaine = date_obj.isocalendar().week
                                nouvelle_ligne[champ_cible] = f"{chantier} {annee} S{semaine + 21}"
                            except Exception as e:
                                print(f"❌ Erreur génération intitulé_rapport : {e}")
                                nouvelle_ligne[champ_cible] = ""

                        elif source == "head":
                            # récupère la valeur dans la ligne_entete correspondant à la colonne col
                            try:
                                idx = df.columns.get_loc(col)
                                nouvelle_ligne[champ_cible] = ligne_entete.iloc[idx]
                            except Exception:
                                nouvelle_ligne[champ_cible] = col  # fallback
                        elif source == "pk":
                            try:
                                idx = df.columns.get_loc(col)
                                nouvelle_ligne[champ_cible] = pk.iloc[idx]
                            except Exception:
                                nouvelle_ligne[champ_cible] = ""
                        elif source == "site":
                            if chantier == "CO7":
                                try:
                                    idx = df.columns.get_loc(col)
                                    nouvelle_ligne[champ_cible] = site.iloc[idx]
                                except Exception:
                                    nouvelle_ligne[champ_cible] = ""
                            else:
                                nouvelle_ligne[champ_cible] = "NC"
                        elif source == "ouvrage":
                            try:
                                idx = df.columns.get_loc(col)
                                nouvelle_ligne[champ_cible] = ouvrage.iloc[idx]
                            except Exception:
                                nouvelle_ligne[champ_cible] = ""
                        elif source == "0":
                            nouvelle_ligne[champ_cible] = date_obj.year
                        elif source == "1":
                            nouvelle_ligne[champ_cible] = date_obj.month
                        elif source == "2":
                            nouvelle_ligne[champ_cible] = date_obj.isocalendar().week
                        elif source == "s2":
                            week = date_obj.isocalendar().week
                            nouvelle_ligne[champ_cible] = f"S{week}"
                        elif source == "3":
                            nouvelle_ligne[champ_cible] = date_obj.day
                        else:
                            nouvelle_ligne[champ_cible] = source

                    donnees_transformees.append(nouvelle_ligne)

        df_final = pd.DataFrame(donnees_transformees)
        df_final.to_excel(f"{chantier}_BDD_UNIFIEE_{datetime.today().date()}.xlsx", index=False)
        print(f"\n✅ Fichier BDD_UNIFIEE_{chantier} généré ✅ :")

    df_CO6 = get_CO6()
    formatage_data(df_CO6, chantier="CO6")

    df_CO7 = get_CO7()
    formatage_data(df_CO7, chantier="CO7")

def CO5():
    # 📥 Lire toutes les lignes du fichier brut sans header
    for nom_fichier in os.listdir("./data"):
        if "Planning GEME" in nom_fichier and nom_fichier:
            file_path = os.path.join("./data", nom_fichier)
    df_brut = pd.read_excel(file_path, header=0,
                            sheet_name="planning (t) (façon CO11)")
    # 🧠 Extraire la ligne 5 (index 4) pour les types (typo_flux)
    typo_flux_ligne = df_brut.iloc[4]

    # 🏷️ Extraire la ligne 6 (index 5) pour les noms de colonnes
    colonnes = df_brut.iloc[0]

    df = pd.read_excel("./data/Planning GEME CO11 maj 20250604.xlsx", skiprows=6, header=None,
                       sheet_name="planning (t) (façon CO11)")
    df.columns = colonnes

    # 🛠 Colonnes contenant les valeurs numériques à transformer
    colonnes_valeurs = ['Cl1', 'Cl1s', 'Cl2', 'Cl3b']  # ⚠️ Attention à la casse exacte dans le fichier

    # 🔁 Charger le mapping existant (sans le modifier comme demandé)
    mapping_df = pd.read_csv("mapping/CO5_colonnes.csv", dtype=str)
    mapping_dict = dict(zip(mapping_df["Source"], mapping_df["Destination"]))

    # 🔄 Transformation
    donnees_transformees = []
    today = datetime.today().strftime("%d/%m/%Y")

    for _, row in df.iterrows():
        date_val = row.iloc[3]
        if pd.isna(date_val):
            continue

        for col in colonnes_valeurs:
            if col not in df.columns:
                continue
            valeur = row[col]
            if pd.notna(valeur) and isinstance(valeur, (int, float)):
                nouvelle_ligne = {}
                for champ_cible, source in mapping_dict.items():
                    source = str(source).strip().lower()
                    if source == "valeur":
                        nouvelle_ligne[champ_cible] = valeur
                    if source == "co5":
                        nouvelle_ligne[champ_cible] = "CO5"
                    elif source == "today":
                        nouvelle_ligne[champ_cible] = today
                    elif source == "intitulé_rapport":
                        annee = ""
                        semaine = ""

                        # Année depuis colonne 0
                        try:
                            val_annee = row.iloc[0]
                            if pd.notna(val_annee):
                                annee = str(int(float(val_annee)))
                        except:
                            annee = ""

                        # Semaine depuis colonne 2 + 1 (ou pas selon besoin)
                        try:
                            val_semaine = row.iloc[2]
                            if pd.notna(val_semaine):
                                semaine_clean = re.sub(r"\D", "", str(val_semaine))
                                semaine = semaine_clean
                        except:
                            semaine = ""

                        nouvelle_ligne[champ_cible] = f"CO5 {annee} S{semaine}".strip()

                    elif source == "0":  # Année
                        try:
                            annee = int(row.iloc[0])
                            nouvelle_ligne[champ_cible] = str(annee)
                        except:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "1":  # Mois
                        try:
                            mois = int(pd.to_datetime(row.iloc[3]).month)
                            nouvelle_ligne[champ_cible] = str(mois).zfill(2)
                        except:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "2":  # Semaine
                        semaine = row.iloc[2]
                        if pd.notna(semaine):
                            semaine_clean = re.sub(r"\D", "", str(semaine))
                            nouvelle_ligne[champ_cible] = semaine_clean
                        else:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "3":  # Jour
                        try:
                            jour = int(pd.to_datetime(row.iloc[3]).day)
                            nouvelle_ligne[champ_cible] = str(jour).zfill(2)
                        except:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "s2":  # Semaine avec S devant
                        semaine = row.iloc[2]
                        if pd.notna(semaine):
                            match = re.search(r'\d+', str(semaine))
                            if match:
                                nouvelle_ligne[champ_cible] = f"S{match.group()}"
                            else:
                                nouvelle_ligne[champ_cible] = ""
                        else:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "head":
                        nouvelle_ligne[champ_cible] = col
                    elif source == "":
                        nouvelle_ligne[champ_cible] = ""
                    else:
                        nouvelle_ligne[champ_cible] = source  # NC, texte fixe etc.

                donnees_transformees.append(nouvelle_ligne)

    # ✅ Créer le DataFrame final
    df_final = pd.DataFrame(donnees_transformees)

    # 💾 Export
    df_final.to_excel(f"CO5_BDD_UNIFIEE_{datetime.today().date()}.xlsx", index=False)

    print("✅ Fichier BDD_UNIFIEE_CO5 généré ✅")

# Création de la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Interface Fonctions")
fenetre.geometry("400x300")

# Boutons
tk.Button(fenetre, text="Exécuter CO8", command=CO8).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO67", command=CO67).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO5", command=CO5).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO9", command=CO9).pack(pady=5)

# Zone de texte pour afficher les résultats
zone_texte = tk.Text(fenetre, height=8, width=50)
zone_texte.pack(pady=10)

# Redirection du print vers la zone de texte
sys.stdout = RedirectText(zone_texte)

# Boucle principale
fenetre.mainloop()