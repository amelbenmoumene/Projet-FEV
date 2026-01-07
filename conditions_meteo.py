import pandas as pd

def preparer_meteo_bordeaux():
    df = pd.read_csv(
        "data/observation-meteorologique-historiques-bordeaux-metropole-synop.csv",
        sep=";",
        encoding="utf-8",
        low_memory=False
    )

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", utc=True)
    df = df[df["Date"].dt.year.between(2017, 2023)]

    df = df.rename(columns={
        "Précipitations dans la dernière heure": "precip_1h",
        "Vitesse du vent moyen 10 mn": "wind_speed",
        "Température (°C)": "temperature_c",
        "Visibilité horizontale": "visibility"
    })

    # Définition condition météo
    def definir_condition(row):
        if pd.notna(row["precip_1h"]) and row["precip_1h"] > 0:
            return "pluie"
        elif pd.notna(row["wind_speed"]) and row["wind_speed"] >= 8:
            return "vent_fort"
        elif pd.notna(row["temperature_c"]) and row["temperature_c"] >= 30:
            return "forte_chaleur"
        elif pd.notna(row["visibility"]) and row["visibility"] >= 20000:
            return "ensoleille"
        else:
            return "normal"

    df["condition_meteo"] = df.apply(definir_condition, axis=1)

    conditions_utiles = [
        "pluie",
        "vent_fort",
        "ensoleille",
        "forte_chaleur"
    ]

    df = df[df["condition_meteo"].isin(conditions_utiles)]
    df_final = df[[
        "Date",
        "condition_meteo",
        "temperature_c",
        "wind_speed",
        "precip_1h",
        "visibility"
    ]]

    df_final.to_csv(
        "meteo_conditions_bordeaux.csv",
        index=False,
        encoding="utf-8-sig"
    )

    return df_final
df_meteo = preparer_meteo_bordeaux()
