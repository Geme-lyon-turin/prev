import os
from calendar import monthrange

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
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
            if col == "TOTAL":
                continue

            try:
                valeur_num = float(str(valeur_brute).replace(",", "."))
                if valeur_num > 0:
                    nouvelle_ligne = {}

                    for champ_cible, source in mapping_dict.items():
                        source_str = str(source).strip()

                        if source_str == "valeur":
                            nouvelle_ligne[champ_cible] = f"{valeur_num:.2f}".replace(".", ",")
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
                        elif source_str == "Jour":
                            try:
                                if "Jour" in row and pd.notna(row["Jour"]):
                                    jour_date = pd.to_datetime(row["Jour"], dayfirst=True, errors="coerce")
                                    if pd.isna(jour_date):
                                        nouvelle_ligne[champ_cible] = ""
                                    else:
                                        excel_epoch = datetime(1899, 12, 30)
                                        jours_excel = (jour_date - excel_epoch).days
                                        nouvelle_ligne[champ_cible] = jours_excel
                                else:
                                    nouvelle_ligne[champ_cible] = ""
                            except Exception as e:
                                print(f"⛔ Erreur conversion Excel numérique 'Jour' : {e}")
                                nouvelle_ligne[champ_cible] = ""
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
    df_mode = df_brut.iloc[5]  # ✅ ligne avec les infos de mode d'excavation et d'évacuation

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
            if pd.notna(valeur) and isinstance(valeur, (int, float)) and valeur > 0:
                nouvelle_ligne = {}
                numero_colonne = df.columns.get_loc(col)
                for champ_cible, source in mapping_dict.items():
                    source_str = str(source).strip()

                    # Normalisation du nom de colonne
                    col_base = "0/4" if col == "0/4.1" else col

                    if source_str == "valeur":
                        nouvelle_ligne[champ_cible] = f"{valeur:.2f}".replace(".", ",")
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
                        # Aller chercher la valeur de head-1 (comme dans "prod")
                        i = 0
                        valeur_typo = ""
                        while (numero_colonne - i) >= 0:
                            candidat = typo_flux_ligne[numero_colonne - i]
                            if pd.notna(candidat) and str(candidat).strip() != "":
                                valeur_typo = candidat
                                break
                            i += 1

                        # Condition : si head-1 == "Production MATEX (tonnes)"
                        if str(valeur_typo).strip() == "Production MATEX (tonnes)":
                            nouvelle_ligne[champ_cible] = "CL2"
                        else:
                            if col_base == "CL1.1" or col_base == "CL1.2" or col_base == "CL1.3" or col_base == "CL1.4":
                                col_base = "CL1"
                            if col_base == "CL2.1" or col_base == "CL2.2" or col_base == "CL2.3" or col_base == "CL2.4":
                                col_base = "CL2"
                            if col_base == "0/4.1" or col_base == "CL2.2" or col_base == "CL2.3" or col_base == "CL2.4":
                                col_base = "0/4"
                            if col_base == "CL3a.1" or col_base == "CL3a.2":
                                col_base = "CL3a"
                            if col_base == "CL3b.1" or col_base == "CL3b.2":
                                col_base = "CL3b"
                            nouvelle_ligne[champ_cible] = col_base
                    elif source_str == "prod":
                        i = 0
                        valeur_typo = ""
                        while (numero_colonne - i) >= 0:
                            candidat = typo_flux_ligne[numero_colonne - i]
                            if pd.notna(candidat) and str(candidat).strip() != "":
                                valeur_typo = candidat
                                break
                            i += 1

                            # Si head-1 vaut "Production MATEX (tonnes)", on garde col_base
                        if str(valeur_typo).strip() == "Production MATEX (tonnes)":
                            nouvelle_ligne[champ_cible] = col_base
                        else:
                            nouvelle_ligne[champ_cible] = "NC"
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
                            # Tente de convertir en datetime
                            if isinstance(val, (pd.Timestamp, datetime)):
                                date_obj = val
                            else:
                                date_obj = pd.to_datetime(val, dayfirst=True, errors="coerce")

                            if pd.isna(date_obj):
                                nouvelle_ligne[champ_cible] = ""
                            else:
                                excel_epoch = datetime(1899, 12, 30)
                                jours_excel = (date_obj - excel_epoch).days
                                nouvelle_ligne[champ_cible] = jours_excel
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
                        # Trouver l'index du premier caractère ')'
                        index = valeur_typo.find(')')
                        # Si un caractère ')' est trouvé, tronquer la chaîne
                        if index != -1:
                            valeur_typo = valeur_typo[:index + 1]
                        nouvelle_ligne[champ_cible] = valeur_typo
                    elif source_str == "evac":
                        try:
                            val_mode = df_mode[numero_colonne]
                            if pd.notna(val_mode) and str(val_mode).strip() != "":
                                nouvelle_ligne[champ_cible] = str(val_mode).strip()
                            else:
                                nouvelle_ligne[champ_cible] = "NC"
                        except:
                            nouvelle_ligne[champ_cible] = "NC"
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
        ligne_entete = raw_header.iloc[5]
        pk = raw_header.iloc[4]
        zone = raw_header.iloc[3]  # Ajout ici ✅
        site = raw_header.iloc[1]
        ouvrage = raw_header.iloc[2]
        return ligne_entete, pk, site, ouvrage, zone

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

        ligne_entete, pk, site, ouvrage, zone = get_raw_header(chantier)

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
                            nouvelle_ligne[champ_cible] = f"{valeur:.2f}".replace(".", ",")
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
                        elif source == "zone":
                            try:
                                idx = df.columns.get_loc(col)
                                nouvelle_ligne[champ_cible] = zone.iloc[idx]  # ✅ Remplace site par zone
                            except Exception:
                                nouvelle_ligne[champ_cible] = ""
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
                            try:
                                if isinstance(date_obj, (pd.Timestamp, datetime)):
                                    excel_epoch = datetime(1899, 12, 30)
                                    jours_excel = (date_obj - excel_epoch).days
                                    nouvelle_ligne[champ_cible] = jours_excel
                                else:
                                    nouvelle_ligne[champ_cible] = ""
                            except Exception:
                                nouvelle_ligne[champ_cible] = ""
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
    df_brut = pd.read_excel(file_path, header=0, sheet_name="planning (t) (façon CO11)")

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

    mois_fr_map = {
        "janvier": 1,"janv": 1,"jan": 1, "février": 2,"fév": 2,"fev": 2, "mars": 3, "avril": 4,"avr": 4,"av": 4,
        "mai": 5, "juin": 6, "juil": 7,"juillet": 7, "août": 8,
        "septembre": 9, "sept":9, "octobre": 10, "oct":10, "novembre": 11,"nov": 11, "décembre": 12,"déc": 12,"dec": 12
    }

    def obtenir_date_excel(jour, mois, annee):
        """
        Renvoie la date Excel (numérique) en soustrayant 1 jour à une date donnée.
        Si le jour vaut 1, recule d'un mois et prend le dernier jour du mois précédent.
        """
        try:
            jour = int(jour)
            mois = int(mois)
            annee = int(annee)

            if jour > 1:
                date_corrigee = datetime(annee, mois, jour - 1)
            else:
                if mois == 1:
                    mois = 12
                    annee -= 1
                else:
                    mois -= 1
                dernier_jour = monthrange(annee, mois)[1]
                date_corrigee = datetime(annee, mois, dernier_jour)

            excel_epoch = datetime(1899, 12, 30)
            return (date_corrigee - excel_epoch).days

        except Exception as e:
            print(f"⛔ Erreur dans 'obtenir_date_excel' : {e}")
            return ""

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
                    if source == "valeur":
                        nouvelle_ligne[champ_cible] = f"{valeur:.2f}".replace(".", ",")
                    elif source == "co5":
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
                            mois_str = str(row.iloc[1]).strip().lower()
                            mois_num = mois_fr_map.get(mois_str, None)
                            if mois_num:
                                nouvelle_ligne[champ_cible] = str(mois_num).zfill(2)
                            else:
                                nouvelle_ligne[champ_cible] = ""
                        except:
                            nouvelle_ligne[champ_cible] = ""

                    elif source == "2":  # Semaine
                        semaine = row.iloc[2]
                        if pd.notna(semaine):
                            semaine_clean = re.sub(r"\D", "", str(semaine))
                            nouvelle_ligne[champ_cible] = semaine_clean
                        else:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "3":
                        try:
                            date_val = pd.to_datetime(row.iloc[3], dayfirst=True, errors='coerce')
                            if pd.isna(date_val):
                                raise ValueError(f"Date invalide: {row.iloc[3]}")
                            jour = date_val.day
                            mois_str = str(row.iloc[1]).strip().lower()
                            mois = mois_fr_map.get(mois_str)
                            annee = int(float(row.iloc[0]))
                            if mois and jour and annee:
                                jours_excel = obtenir_date_excel(jour, mois, annee)
                                nouvelle_ligne[champ_cible] = jours_excel
                            else:
                                nouvelle_ligne[champ_cible] = ""
                        except Exception as e:
                            print(f"⛔ Erreur date complète (Excel) : {e}")
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

def CO6_new(chantier="CO67", mapping_path="mapping/CO6_new_colonnes.csv"):
    # 🔍 Cherche le bon fichier
    for nom_fichier in os.listdir("./data"):
        if "TELT_CO67" in nom_fichier and "lock" not in nom_fichier:
            file_path = os.path.join("./data", nom_fichier)
            break
    else:
        print("❌ Fichier CO67 introuvable.")
        return

    print(f"📄 Fichier trouvé : {file_path}")

    # 🔍 Lecture brute pour entêtes utiles
    df_raw = pd.read_excel(file_path, sheet_name="Prévision_hebdo", header=None)
    ligne_cession = df_raw.iloc[2]
    ligne_mode = df_raw.iloc[5]
    ligne_formation = df_raw.iloc[7]
    ligne_ouvrage = df_raw.iloc[6]
    ligne_pk = df_raw.iloc[8]
    ligne_head = df_raw.iloc[9]

    # 🧠 Lecture des vraies données
    df = pd.read_excel(file_path, sheet_name="Prévision_hebdo", header=9)

    # 📅 Création colonne Date à partir d'année, mois, jour
    try:
        df["Date"] = pd.to_datetime(
            df.iloc[:, 1].astype(str) + "-" +
            df.iloc[:, 2].astype(str).str.zfill(2) + "-" +
            df.iloc[:, 4].astype(str).str.extract(r'(\d{1,2})')[0].str.zfill(2),
            format="%Y-%m-%d",  # <- Ajout du format explicite
            errors="coerce"
        )
        df["Date"] = df["Date"].dt.date  # Pour enlever l'heure
    except Exception as e:
        print(f"❌ Erreur génération date : {e}")
        return

    print("✅ Colonne 'Date' générée.")

    # 🗂️ Chargement du mapping
    mapping_df = pd.read_csv(mapping_path)
    mapping_dict = dict(zip(mapping_df["Destination"], mapping_df["Source"]))

    colonnes_valeurs = [
        'Cl1', 'Cl2', 'Cl3', 'Cl1s', 'cl3a', 'Cl3b',
        'sable 0/4mm', 'granulats 4/8mm', 'granulats 8/16mm',
        'sable 0/4CLmm', 'granulats 10/20mm',
        'CR', 'Fiche n°7 IG9091 SNCF', 'Fiche n°9 IG9091 SNCF',
        "SETRA Note d'information n°34", 'CDF', 'PST IG 90260',
        'ZH simplifiée C0 Phi30°'
    ]  # À adapter selon ton CSV mapping

    colonnes_utiles = [col for col in df.columns if any(val.lower() in col.lower() for val in colonnes_valeurs)]

    today = datetime.today().strftime("%d/%m/%Y")
    donnees_finales = []

    # Exemple dans ta boucle principale
    for i, row in df.iterrows():
        date_obj = row["Date"]
        if pd.isna(row.iloc[1]):  # Si colonne Année vide, on saute
            continue
        colonnes_valeurs_trouvees = [col for col in colonnes_utiles if pd.notna(row[col])]
        col_valeur = colonnes_valeurs_trouvees[0] if colonnes_valeurs_trouvees else None
        idx_valeur = df.columns.get_loc(col_valeur) if col_valeur else None
        for col in colonnes_utiles:
            valeur = row[col]
            if pd.notna(valeur) and valeur > 0:
                idx_col = df.columns.get_loc(col)
                nouvelle_ligne = {}
                for champ_cible, source in mapping_dict.items():
                    if source == "valeur":
                        try:
                            val_float = float(valeur)
                            val_formate = f"{val_float:.5f}".replace(".", ",")
                            val_formate = val_formate.rstrip('0').rstrip(',') if ',' in val_formate else val_formate
                            nouvelle_ligne[champ_cible] = val_formate
                        except:
                            nouvelle_ligne[champ_cible] = str(valeur).replace(".", ",")

                    elif source == "Chantier":
                        nouvelle_ligne[champ_cible] = chantier
                    elif source == "today":
                        nouvelle_ligne[champ_cible] = today
                    elif source == "intitulé_rapport":
                        if pd.isna(date_obj):
                            nouvelle_ligne[champ_cible] = ""
                        else:
                            try:
                                # S’il s’agit d’un datetime.date ou datetime.datetime, isocalendar() renvoie un tuple (année, semaine, jour)
                                semaine = date_obj.isocalendar()[1]  # index 1 = semaine ISO
                                annee = date_obj.year
                                nouvelle_ligne[champ_cible] = f"{chantier} {annee} S{semaine}"
                            except Exception as e:
                                print(f"⚠️ Erreur intitulé_rapport : {e} avec date {date_obj}")
                                nouvelle_ligne[champ_cible] = ""
                    elif source == "head":
                        if idx_col is not None:
                            nouvelle_ligne[champ_cible] = ligne_head.iloc[idx_col]
                        else:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "Intitulé_rapport":
                        try:
                            annee = date_obj.year
                            semaine = date_obj.isocalendar().week
                            nouvelle_ligne[champ_cible] = f"{chantier} {annee} S{semaine}"
                        except:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "pk":
                        idx = df.columns.get_loc(col)
                        nouvelle_ligne[champ_cible] = ligne_pk.iloc[idx]
                    elif source == "evac":
                        idx = df.columns.get_loc(col)
                        nouvelle_ligne[champ_cible] = ligne_mode.iloc[idx]
                    elif source == "formation":
                        idx = df.columns.get_loc(col)
                        nouvelle_ligne[champ_cible] = ligne_formation.iloc[idx]

                    elif source == "NC":
                        nouvelle_ligne[champ_cible] = "NC"
                    elif source == "S3":
                        val = row.iloc[3]
                        if pd.isna(val):
                            nouvelle_ligne[champ_cible] = ""
                        else:
                            nouvelle_ligne[champ_cible] = f"S{int(val)}"
                    elif source == "cession":
                        if idx_col is not None:
                            col_letter = get_column_letter(idx_col + 1)  # +1 car Excel commence à 1
                            col_index = column_index_from_string(col_letter)

                            if 8 <= col_index <= 14:  # H (8) à N (14)
                                nouvelle_ligne[champ_cible] = "Production"
                            elif 24 <= col_index <= 54:  # X (24) à BB (54)
                                nouvelle_ligne[champ_cible] = "Cession"
                            elif 56 <= col_index <= 87:  # BD (56) à CI (87)
                                nouvelle_ligne[champ_cible] = "Livraison"
                            else:
                                nouvelle_ligne[champ_cible] = ""
                        else:
                            nouvelle_ligne[champ_cible] = ""
                    elif source == "Code_SITE":
                        code_site = ""
                        noText = ["Point de production","Point de cession","Point de livraison"]
                        if idx_col is not None:
                            for decalage in range(-4, 3):  # de -4 à +2 inclus
                                index_recherche = idx_col + decalage
                                print(index_recherche)
                                if 0 <= index_recherche < len(ligne_cession):  # s'assurer que l'index est valide
                                    valeur = ligne_cession.iloc[index_recherche]
                                    if pd.notna(valeur) and str(valeur).strip() != "" and valeur not in noText:
                                        code_site = valeur
                                        break
                        nouvelle_ligne[champ_cible] = code_site
                    elif source == "pk":
                        if idx_col is not None:
                            nouvelle_ligne[champ_cible] = ligne_pk.iloc[idx_col]
                        else:
                            nouvelle_ligne[champ_cible] = ""
                    elif "ouvrage" in source:
                        if idx_col is not None:
                            # Récupère la valeur de la cellule, qui contient plusieurs lignes séparées par '\n'
                            valeur = ligne_ouvrage.iloc[idx_col]
                            lignes = valeur.split("\n")

                            if "majeur" in source:
                                # Prend la première ligne si elle existe
                                nouvelle_ligne[champ_cible] = lignes[0] if lignes else ""
                            elif "mineur" in source:
                                # Prend la deuxième ligne si elle existe
                                nouvelle_ligne[champ_cible] = lignes[1] if len(lignes) > 1 else ""
                            else:
                                # Si le type n'est ni majeur ni mineur, met une chaîne vide par défaut
                                nouvelle_ligne[champ_cible] = ""
                        else:
                            # Si l'index de la colonne est None, on met vide pour éviter les erreurs
                            nouvelle_ligne[champ_cible] = ""
                    elif source.isdigit():
                        val = row.iloc[int(source)]
                        if source == "4":
                            try:
                                if pd.notna(val):
                                    # Conversion explicite en datetime
                                    date_obj = pd.to_datetime(val, dayfirst=True, errors="coerce")
                                    if pd.isna(date_obj):
                                        raise ValueError(f"Date invalide: {val}")

                                    # Conversion en valeur numérique Excel (base 1899-12-30)
                                    excel_base = datetime(1899, 12, 30)
                                    val = (date_obj - excel_base).days
                                else:
                                    val = ""
                            except Exception as e:
                                print(f"⛔ Erreur date Excel (source == '4') : {e}")
                                val = ""

                        nouvelle_ligne[champ_cible] = val
                    else:
                        nouvelle_ligne[champ_cible] = source

                donnees_finales.append(nouvelle_ligne)

    # 📤 Export
    df_final = pd.DataFrame(donnees_finales)
    nom_fichier = f"CO6NEW_BDD_UNIFIEE_{datetime.today().date()}.xlsx"
    df_final.to_excel(nom_fichier, index=False)
    print(f"\n✅ Fichier généré : {nom_fichier}")


# Création de la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Interface Fonctions")
fenetre.geometry("400x500")

# Boutons
tk.Button(fenetre, text="Exécuter CO5", command=CO5).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO67", command=CO67).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO8", command=CO8).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO9", command=CO9).pack(pady=5)
tk.Button(fenetre, text="Exécuter CO67_new", command=CO6_new).pack(pady=5)

# Zone de texte pour afficher les résultats
zone_texte = tk.Text(fenetre, height=20, width=50)
zone_texte.pack(pady=10)

# Redirection du print vers la zone de texte
sys.stdout = RedirectText(zone_texte)

# Boucle principale
fenetre.mainloop()