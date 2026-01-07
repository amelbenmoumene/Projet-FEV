import pandas as pd
import glob

def fusion_tramway():
    fichiers = glob.glob("data/TRAMWAY-Incidents_20*.xlsx")
    dfs = []

    # Dictionnaire mois français
    mois_fr = {
        1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril', 5: 'mai', 6: 'juin',
        7: 'juillet', 8: 'août', 9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'
    }

    for f in fichiers:
        df = pd.read_excel(f)
        dfs.append(df)
        df.columns = [col.capitalize() for col in df.columns]
        df = df.rename(columns={"Direction": "Voie", "Début": "Debut", "Fin": "Fin"})

    df_final = pd.concat(dfs, ignore_index=True)

    df_final = df_final[df_final['Debut'] != 'Début de service']
    df_final = df_final[df_final['Debut'] != 'Fin de service']

    if 'Rame' in df_final.columns:
        df_final = df_final.drop(columns=['Rame'])

    # Normalisation de la colonne Ligne
    lignes_valides = ['A', 'B', 'C', 'D']
    def normaliser_ligne(val):
        val = str(val)
        for sep in ['/', 'et', ',', '-', ';']:
            val = val.replace(sep, ' ')
        return [c for c in val if c in lignes_valides]

    df_final['Ligne'] = df_final['Ligne'].apply(normaliser_ligne)
    df_final = df_final.explode('Ligne')

    df_final['Debut_dt'] = pd.to_datetime(df_final['Debut'], format='%H:%M:%S', errors='coerce')
    df_final['Fin_dt']   = pd.to_datetime(df_final['Fin'], format='%H:%M:%S', errors='coerce')

    # Calcul de la durée en minutes
    df_final['Duree'] = ((df_final['Fin_dt'] - df_final['Debut_dt']).dt.total_seconds() / 60)
    df_final['Duree'] = df_final['Duree'].fillna(0).astype(int)
    df_final.loc[df_final['Duree'] < 0, 'Duree'] += 24*60  # incident passant minuit

    # Colonne date
    df_final['Date'] = pd.to_datetime(df_final['Date'], errors='coerce', dayfirst=True)
    df_final = df_final.dropna(subset=['Date'])
    df_final = df_final[df_final['Date'].dt.year.between(2017, 2023)]
    df_final['Date'] = df_final['Date'].apply(lambda x: f"{x.day} {mois_fr[x.month]} {x.year}" if pd.notnull(x) else None)

    # Mettre la localisation en minuscule
    if 'Localisation' in df_final.columns:
        df_final['Localisation'] = df_final['Localisation'].astype(str).str.strip().str.lower()

    df_final = df_final.drop(columns=['Debut', 'Fin', 'Debut_dt', 'Fin_dt'], errors='ignore')

    # Supprimer colonnes entièrement vides et lignes incomplètes sur colonnes essentielles
    df_final = df_final.dropna(axis=1, how='all')
    colonnes_obligatoires = ['Date', 'Ligne']
    df_final = df_final.dropna(subset=[col for col in colonnes_obligatoires if col in df_final.columns])

    df_final.to_csv("fusion_tramway.csv", index=False, encoding="utf-8-sig")

    return df_final

df_normalise = fusion_tramway()
print("Valeurs uniques Date :", df_normalise['Date'].unique())
print("Valeurs uniques Ligne après normalisation :", df_normalise['Ligne'].unique())
