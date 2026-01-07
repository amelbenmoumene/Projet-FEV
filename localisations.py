import pandas as pd

def traiter_stops_tramway():

    df = pd.read_csv(
        "data/stops.txt",
        sep=",",
        encoding="utf-8"
    )

    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        "stop_id": "Stop_id",
        "stop_name": "Nom_arret",
        "stop_lat": "Latitude",
        "stop_lon": "Longitude",
        "zone_id": "Zone_id",
        "location_type": "Type_arret",
        "wheelchair_boarding": "Acces_fauteuil"
    })

    # Colonnes utiles uniquement
    colonnes_utiles = [
        "Stop_id",
        "Nom_arret",
        "Latitude",
        "Longitude",
        "Zone_id",
        "Type_arret",
        "Acces_fauteuil"
    ]
    df = df[colonnes_utiles]

    # Suppression des lignes sans coordonnées ou sans nom
    df = df.dropna(subset=["Nom_arret", "Latitude", "Longitude"])
    df["Latitude"] = df["Latitude"].astype(float)
    df["Longitude"] = df["Longitude"].astype(float)

    # Normalisation des types
    df["Type_arret"] = df["Type_arret"].fillna(0).astype(int)
    df["Acces_fauteuil"] = df["Acces_fauteuil"].fillna(0).astype(int)
    df["Nom_arret"] = (df["Nom_arret"].astype(str).str.strip().str.lower())
    # Suppression des doublons (même arrêt géographique)
    df = df.drop_duplicates(subset=["Latitude", "Longitude"])
    df = df.drop_duplicates(subset=["Nom_arret"])

    df.to_csv(
        "stops_tramway_clean.csv",
        index=False,
        encoding="utf-8-sig"
    )

    return df


df_stops = traiter_stops_tramway()
print("Nombre total d'arrêts :", len(df_stops))
print("Exemples d'arrêts :", df_stops["Nom_arret"].unique()[:10])
