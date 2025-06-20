import pandas as pd
from datetime import datetime

# 📂 Charger le fichier brut pour accéder à la ligne 6 (index 5)
df_brut = pd.read_excel("./data/T231117_TELT-CO8_prévision-flux - 2025 S07.xlsm", header=None)

# 🧠 Extraire la ligne 6 (Typo_Flux) AVANT le vrai header
typo_flux_ligne = df_brut.iloc[4]

# 📂 Charger le vrai DataFrame avec header à la ligne 7 (index 6)
df = pd.read_excel("./data/T231117_TELT-CO8_prévision-flux - 2025 S07.xlsm", header=6)

# 🔁 Charger le mapping des colonnes
mapping_df = pd.read_csv("mapping/CO8_colonnes.csv", dtype=str)
mapping_dict = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))

# ✅ Colonnes avec des valeurs numériques à traiter
colonnes_valeurs = ['0/4', '4/8', '8/16', '0/4.1', '0/4 CL', '4/10', '10/20','eFs', 'eFsg','l1-42','t7SB3','CL1','CL2','CL3a','CL3b','CL1.1','CL2.1','CL3a.1','CL3b.1']

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
                elif source_str == "today":
                    nouvelle_ligne[champ_cible] = today
                elif source_str == "head":
                    nouvelle_ligne[champ_cible] = col_base
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
df_final.to_csv(f"CO8_BDD_UNIFIEE_{datetime.today().date()}.csv", index=False)

print("✅ Fichier BDD_UNIFIEE_CO8 généré ✅")
