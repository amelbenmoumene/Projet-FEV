import pandas as pd
import glob

fichiers = glob.glob("data/TRAMWAY-Incidents_20*.xlsx")


dfs = []
for f in fichiers:
    df = pd.read_excel(f)
    dfs.append(df)
    print(" \n fichiers :", f)
    df.columns = [col.capitalize() for col in df.columns]

    df = df.rename(columns = {"Direction":"Voie"})
    df = df.rename(columns = {"Début":"Heure début"})
    df = df.rename(columns = {"Fin":"Heure fin"})
    for col in df.columns:
        print(col)


 

df_final = pd.concat(dfs, ignore_index = True)
df_final.to_excel("fusion_tramway.xlsx", index = False)