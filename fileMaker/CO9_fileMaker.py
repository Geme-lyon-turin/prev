import pandas as pd
from datetime import datetime
import unicodedata
import re

# 📅 Date actuelle formatée
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
file_path = "./data/Prevision flux MT103 - S06.xlsm"
sheet_name = "Planning besoins t"
print("📥 Chargement du fichier Excel...")

df = pd.read_excel(file_path, header=13, sheet_name=sheet_name)
df.columns = [nettoyer_texte(col) for col in df.columns]  # Nettoyage des noms de colonnes

print("✅ Données chargées avec succès !")
print("🔎 Colonnes disponibles dans le fichier Excel :")
for col in df.columns:
    print(f" - {col}")

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
    else:
        print(f"⚠️ Colonne absente (après nettoyage/normalisation) : {col_norm}")

# 📥 Chargement du mapping des champs de sortie
print("📥 Chargement du mapping CO9_colonnes.csv...")
mapping_df = pd.read_csv("mapping/CO9_colonnes.csv", dtype=str)
mapping_dict = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))

print("✅ Mapping chargé. Aperçu :")
print(mapping_dict)

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
                    elif source_str in ["head", "colonne_name"]:
                        nouvelle_ligne[champ_cible] = col
                    elif source_str == "Typo courte":
                        col_normalise = normaliser_texte(col)
                        nouvelle_ligne[champ_cible] = typo_courte_dict.get(col_normalise, "")
                        if not nouvelle_ligne[champ_cible]:
                            print(
                                f"⚠️ Aucun match trouvé pour '{col}' (normalisé: '{col_normalise}') dans Typo courte.")

                    elif source_str.isdigit():
                        try:
                            nouvelle_ligne[champ_cible] = str(int(row.iloc[int(source_str)]))
                        except Exception:
                            nouvelle_ligne[champ_cible] = ""
                    else:
                        nouvelle_ligne[champ_cible] = source_str

                donnees_transformees.append(nouvelle_ligne)
                #print(f"✅ Ligne ajoutée : {nouvelle_ligne}")

        except Exception as e:
            print(f"❌ Impossible de convertir la valeur : {valeur_brute} → {e}")
            continue

# 📤 Export CSV
df_final = pd.DataFrame(donnees_transformees)

output_file = f"CO9_BDD_UNIFIEE_{datetime.today().date()}.csv"
df_final.to_csv(output_file, index=False)

print("\n✅ Export terminé ! Fichier généré :", output_file)
print("🧾 Aperçu du DataFrame final :")
print(df_final.head())
