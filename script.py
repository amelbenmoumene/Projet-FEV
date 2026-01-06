import pandas as pd
import glob
import re

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
        df = df.rename(columns={"Direction": "Voie", "Début": "Heure début", "Fin": "Heure fin"})

    df_final = pd.concat(dfs, ignore_index=True)

    # Supprimer les lignes de début/fin de service
    df_final = df_final[df_final['Debut'] != 'Début de service']
    df_final = df_final[df_final['Debut'] != 'Fin de service']

    # Supprimer la colonne Rame
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
    
    # Filtrer Debut et Fin au format HH:MM:SS
    pattern_heure = r'^\d{2}:\d{2}:\d{2}$'
    df_final['Debut'] = df_final['Debut'].astype(str).str.strip()
    df_final = df_final[df_final['Debut'].str.match(pattern_heure, na=False)]
    df_final['Fin'] = df_final['Fin'].astype(str).str.strip()
    df_final = df_final[df_final['Fin'].str.match(pattern_heure, na=False)]

    colonnes_obligatoires = ['Date', 'Debut', 'Fin', 'Ligne']

    # Conversion Debut et Fin en datetime
    df_final['Debut'] = pd.to_datetime(df_final['Debut'], format='%H:%M:%S', errors='coerce')
    df_final['Fin']   = pd.to_datetime(df_final['Fin'], format='%H:%M:%S', errors='coerce')

    # Calcul de la durée en minutes (entier)
    df_final['Duree'] = ((df_final['Fin'] - df_final['Debut']).dt.total_seconds() / 60).astype(int)
    df_final.loc[df_final['Duree'] < 0, 'Duree'] += 24*60  # Incident passe minuit

    # Filtrage Date au format strict YYYY-MM-DD 00:00:00
    pattern_date = r'^\d{4}-\d{2}-\d{2} 00:00:00$'
    df_final['Date'] = df_final['Date'].astype(str).str.strip()
    df_final = df_final[df_final['Date'].str.match(pattern_date, na=False)]

    # Conversion Date en mois + année français
    df_final['Date'] = pd.to_datetime(df_final['Date'], errors='coerce')
    df_final['Date'] = df_final['Date'].apply(lambda x: f"{mois_fr[x.month]} {x.year}" if pd.notnull(x) else None)

    # Supprimer colonnes inutiles pour Power BI
    df_final = df_final.drop(columns=['Debut', 'Fin'])
    
    # Supprimer colonnes entièrement vides
    df_final = df_final.dropna(axis=1, how='all')
    
    # Supprimer lignes incomplètes
    df_final = df_final.dropna(subset=[col for col in colonnes_obligatoires if col in df_final.columns])
    
    # Export CSV
    df_final.to_csv("fusion_tramway.csv", index=False, encoding="utf-8-sig")

    return df_final

df_normalise = fusion_tramway()
print("Valeurs uniques Date :", df_normalise['Date'].unique())
print("Valeurs uniques Ligne après normalisation :", df_normalise['Ligne'].unique())
