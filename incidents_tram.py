import pandas as pd
import glob
import re

def fusion_tramway_csv():
    fichiers = glob.glob("data/TRAMWAY-Incidents_20*.xlsx")
    dfs = []

    # Dictionnaire mois français
    mois_fr = {
        1:'janvier',2:'février',3:'mars',4:'avril',
        5:'mai',6:'juin',7:'juillet',8:'août',
        9:'septembre',10:'octobre',11:'novembre',12:'décembre'
    }

    for f in fichiers:
        df = pd.read_excel(f)

        df.columns = [col.capitalize() for col in df.columns]
        df = df.rename(columns={
            "Direction":"Voie",
            "Début":"Debut",
            "Fin":"Fin"
        })

        dfs.append(df)

    df_final = pd.concat(dfs, ignore_index=True)

    # Suppression lignes inutiles
    df_final = df_final[~df_final['Debut'].isin(['Début de service','Fin de service'])]

    if 'Rame' in df_final.columns:
        df_final = df_final.drop(columns=['Rame'])

    # Normalisation Ligne
    lignes_valides = ['A','B','C','D']
    def normaliser_ligne(val):
        val = str(val)
        for sep in ['/','et',',','-',';']:
            val = val.replace(sep,' ')
        return [c for c in val if c in lignes_valides]

    df_final['Ligne'] = df_final['Ligne'].apply(normaliser_ligne)
    df_final = df_final.explode('Ligne')

    # Durée
    df_final['Debut_dt'] = pd.to_datetime(df_final['Debut'], errors='coerce')
    df_final['Fin_dt'] = pd.to_datetime(df_final['Fin'], errors='coerce')
    df_final['Duree'] = ((df_final['Fin_dt'] - df_final['Debut_dt']).dt.total_seconds() / 60)
    df_final['Duree'] = df_final['Duree'].fillna(0).astype(int)
    df_final.loc[df_final['Duree'] < 0, 'Duree'] += 1440

    # ---- Gestion Date propre ----
    def convertir_date(val):
        try:
            d = pd.to_datetime(val, dayfirst=True)
            return f"{d.day} {mois_fr[d.month]} {d.year}"
        except:
            return None

    df_final['Date'] = df_final['Date'].apply(convertir_date)

    # Regex format : "1 janvier 2017"
    pattern = r'^\d{1,2} (janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre) \d{4}$'

    # Supprimer lignes invalides
    df_final = df_final[df_final['Date'].astype(str).str.match(pattern)]

    # Localisation
    if 'Localisation' in df_final.columns:
        df_final['Localisation'] = df_final['Localisation'].astype(str).str.strip().str.lower()

    # Nettoyage final
    df_final = df_final.drop(columns=['Debut','Fin','Debut_dt','Fin_dt'], errors='ignore')

    # Export CSV
    df_final.to_csv("fusion_tramway.csv", index=False, encoding='utf-8-sig')

    return df_final

# Exécution
df = fusion_tramway_csv()
print("✔ Fichier CSV prêt pour Power BI")
print("Exemple Date :", df['Date'].head())
