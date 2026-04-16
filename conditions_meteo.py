import pandas as pd
import numpy as np

df = pd.read_excel(
    "data/open-meteo-44.82N0.56W14m.xlsx",
    skiprows=3
)

modiff

df.columns = (
    df.columns
    .str.lower()
    .str.strip()
    .str.replace(" ", "_")
    .str.replace("(", "", regex=False)
    .str.replace(")", "", regex=False)
    .str.replace("°c", "c")
    .str.replace("%", "pct")
)

if "time" not in df.columns:
    raise ValueError("La colonne 'time' n'a pas été trouvée")

df["time"] = pd.to_datetime(df["time"], errors="coerce")

colonnes_numeriques = [
    "temperature_2m_c",
    "apparent_temperature_c",
    "relative_humidity_2m_pct",
    "wind_speed_10m_km/h",
    "wind_gusts_10m_km/h",
    "precipitation_mm",
    "rain_mm"
]

for col in colonnes_numeriques:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
df = df.dropna(subset=["time"])

df = df.sort_values("time")
df = df.set_index("time")

for col in colonnes_numeriques:
    if col in df.columns:
        df[col] = df[col].interpolate(method="time")

df = df.reset_index()

def definir_condition_meteo(row):
    temp = row["temperature_2m_c"]
    pluie = row["precipitation_mm"]
    vent = row["wind_gusts_10m_km/h"]
    humidite = row["relative_humidity_2m_pct"]

    if pluie >= 10:
        return "Pluie intense"
    elif pluie > 0:
        return "Pluie"
    elif temp <= 0:
        return "Gel"
    elif temp >= 30:
        return "Forte chaleur"
    elif vent >= 50:
        return "Vent fort"
    elif humidite >= 95 and temp <= 5:
        return "Brouillard"
    else:
        return "Temps calme"

df["condition_meteo"] = df.apply(definir_condition_meteo, axis=1)

colonnes_finales = [
    "time",
    "temperature_2m_c",
    "relative_humidity_2m_pct",
    "wind_speed_10m_km/h",
    "wind_gusts_10m_km/h",
    "precipitation_mm",
    "condition_meteo"
]

df = df[colonnes_finales]

df.to_csv(
    "meteo_bordeaux_cleaned.csv",
    sep=";",
    decimal=",",
    index=False
)

print("Script exécuté avec succès")
print("Une seule colonne condition_meteo, données propres")
