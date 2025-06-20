import pandas as pd
from datetime import datetime

# 📥 Lire toutes les lignes du fichier brut sans header
df_brut = pd.read_excel("./data/Planning GEME CO11 maj 20250604.xlsx", header=0, sheet_name="planning (t) (façon CO11)")
# 🧠 Extraire la ligne 5 (index 4) pour les types (typo_flux)
typo_flux_ligne = df_brut.iloc[4]

# 🏷️ Extraire la ligne 6 (index 5) pour les noms de colonnes
colonnes = df_brut.iloc[0]

df = pd.read_excel("./data/Planning GEME CO11 maj 20250604.xlsx", skiprows=6, header=None, sheet_name="planning (t) (façon CO11)")
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
        continue  # 🧼 Ligne ignorée car la date (colonne 3) est vide

    for col in colonnes_valeurs:
        if col not in df.columns:
            continue
        valeur = row[col]
        if pd.notna(valeur) and isinstance(valeur, (int, float)):
            nouvelle_ligne = {}
            col_index = df.columns.get_loc(col)
            for champ_cible, source in mapping_dict.items():
                source = str(source).strip()
                if source == "valeur":
                    nouvelle_ligne[champ_cible] = valeur
                elif source == "today":
                    nouvelle_ligne[champ_cible] = today
                elif source == "3":
                    date_val = row.iloc[3]
                    if pd.notna(date_val):
                        date_val = pd.to_datetime(date_val)
                        nouvelle_ligne[champ_cible] = date_val.strftime("%d")
                    else:
                        continue

                elif source == "head":
                    nouvelle_ligne[champ_cible] = col
                elif source.isdigit():
                    val = row.iloc[int(source)]
                    nouvelle_ligne[champ_cible] = date_val.strftime("%d")
                elif source == "":
                    nouvelle_ligne[champ_cible] = ""
                else:
                    nouvelle_ligne[champ_cible] = source  # Texte brut (NC, CO8, etc.)
            donnees_transformees.append(nouvelle_ligne)

# ✅ Créer le DataFrame final
df_final = pd.DataFrame(donnees_transformees)

# 💾 Export
df_final.to_csv(f"CO5_BDD_UNIFIEE_{datetime.today().date()}.csv", index=False)

print("✅ Fichier BDD_UNIFIEE_CO8 généré ✅")
