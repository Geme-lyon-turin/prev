import pandas as pd
from datetime import datetime

def get_raw_header(chantier):
    raw_header = pd.read_excel("./data/CO67_S25-04.xlsx", header=None, sheet_name=chantier)
    # lignes importantes
    ligne_entete = raw_header.iloc[5]  # index 5 = ligne 6 Excel
    pk = raw_header.iloc[4]             # index 3 = ligne 4 Excel
    site = raw_header.iloc[1]           # index 0 = ligne 1 Excel
    ouvrage = raw_header.iloc[2]        # index 1 = ligne 2 Excel
    return ligne_entete, pk, site, ouvrage

def get_CO6():
    df = pd.read_excel("./data/CO67_S25-04.xlsx", header=6, sheet_name="CO6")
    df['Jour'] = df["Pré classement\nCorrigé à front"].astype(str).str.strip()
    df['Date'] = pd.to_datetime(df["Unnamed: 1"], errors='coerce')
    df = df[df['Date'].notna()]
    return df

def get_CO7():
    df = pd.read_excel("./data/CO67_S25-04.xlsx", header=6, sheet_name="CO7")
    df['Jour'] = df["Pré classement\nCorrigé à front"].astype(str).str.strip()
    df['Date'] = pd.to_datetime(df["Unnamed: 1"], errors='coerce')
    df = df[df['Date'].notna()]
    return df

def formatage_data(df, chantier):
    colonnes_valeurs = ['cl1','cl2', 'cl3']  # à adapter selon ton cas
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
                            nouvelle_ligne[champ_cible] = f"{chantier} {annee} S{semaine+21}"
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
                    elif source == "2+21":
                        week = date_obj.isocalendar().week
                        nouvelle_ligne[champ_cible] = f"S{week + 21}"
                    elif source == "3":
                        nouvelle_ligne[champ_cible] = date_obj.day
                    else:
                        nouvelle_ligne[champ_cible] = source

                donnees_transformees.append(nouvelle_ligne)

    df_final = pd.DataFrame(donnees_transformees)
    df_final.to_csv(f"{chantier}_BDD_UNIFIEE_{datetime.today().date()}.csv", index=False)
    print(f"\n✅ Export terminé pour {chantier} :")
    print(df_final.head())

df_CO6 = get_CO6()
formatage_data(df_CO6, chantier="CO6")

df_CO7 = get_CO7()
formatage_data(df_CO7, chantier="CO7")
