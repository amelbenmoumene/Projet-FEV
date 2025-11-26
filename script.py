import pandas as pd
import glob

def fusion_tramway():
    fichiers = glob.glob("data/TRAMWAY-Incidents_20*.xlsx")

    dfs = []
    for f in fichiers:
        df = pd.read_excel(f)
        dfs.append(df)
        df.columns = [col.capitalize() for col in df.columns]

        df = df.rename(columns = {"Direction":"Voie"})
        df = df.rename(columns = {"Début":"Heure début"})
        df = df.rename(columns = {"Fin":"Heure fin"})

    df_final = pd.concat(dfs, ignore_index = True)
    df_final = df_final[df_final['Debut'] != 'Début de service']
    df_final = df_final[df_final['Debut'] != 'Fin de service']

    # Nettoyage Rame
    df_final['Rame'] = df_final['Rame'].astype(str)
    df_final = df_final[df_final['Rame'].notna()]
    df_final = df_final[df_final['Rame'].str.strip() != 'nan']
    df_final = df_final[df_final['Rame'].str.strip() != '']
    df_final = df_final[df_final['Rame'].str.isdigit()]

    # ------------------------
    # Normalisation Ligne
    lignes_valides = ['A', 'B', 'C', 'D']

    def normaliser_ligne(val):
        val = str(val)
        # Remplacer séparateurs communs par espace
        for sep in ['/', 'et', ',', '-', ';']:
            val = val.replace(sep, ' ')
        # Séparer en lignes
        Lignes = []
        for c in val:
            if c in lignes_valides:
                Lignes.append(c)
        return Lignes

    # Appliquer et dupliquer les lignes
    df_final['Ligne'] = df_final['Ligne'].apply(normaliser_ligne)

    # Dupliquer les lignes
    df_final = df_final.explode('Ligne')

    # Sauvegarder le résultat
    df_final.to_excel("fusion_tramway.xlsx", index=False)

    return df_final

df_normalise = fusion_tramway()
print("Valeurs uniques Ligne après normalisation : ", df_normalise['Ligne'].unique())
