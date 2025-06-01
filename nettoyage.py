import pandas as pd

# Charger le fichier Excel
fichier = "crimes-et-delits-enregistres-par-les-services-de-gendarmerie-et-de-police-depuis-2012.xlsx"
excel_file = pd.ExcelFile(fichier)

# Liste des années et des types de services
annees = range(2012, 2022)
services = [("Services PN", "PN"), ("Services GN", "GN")]

# Liste pour stocker les résultats
toutes_annees = []

for annee in annees:
    for prefix, label_service in services:
        sheet_name = f"{prefix} {annee}"
        if sheet_name not in excel_file.sheet_names:
            print(f"Feuille absente : {sheet_name}")
            continue

        df_raw = excel_file.parse(sheet_name, header=None)

        # Extraire les métadonnées de colonnes
        headers_dept = df_raw.iloc[0, 2:]
        headers_perimetre = df_raw.iloc[1, 2:]
        headers_poste = df_raw.iloc[2, 2:]

        # Extraire les types de crimes et les valeurs
        crime_types = df_raw.iloc[3:, 1].reset_index(drop=True)
        crime_values = df_raw.iloc[3:, 2:].reset_index(drop=True)

        # Construire le DataFrame des métadonnées
        meta_info = pd.DataFrame({
            'Département': headers_dept.values,
            'Périmètre': headers_perimetre.values,
            'Poste / Brigade': headers_poste.values
        })

        # Créer les enregistrements
        records = []
        for i, crime in crime_types.items():
            for col in range(crime_values.shape[1]):
                valeur = crime_values.iat[i, col]
                if pd.notnull(valeur) and valeur != 0:
                    records.append({
                        'Année': annee,
                        'Département': meta_info.iloc[col, 0],
                        'Périmètre': meta_info.iloc[col, 1],
                        'Poste / Brigade': meta_info.iloc[col, 2],
                        'Type de crime': crime,
                        'Nombre': valeur,
                        'Service': label_service
                    })

        df_partiel = pd.DataFrame(records)
        toutes_annees.append(df_partiel)

# Fusionner tous les DataFrames
df_final = pd.concat(toutes_annees, ignore_index=True)

# Export CSV
df_final.to_csv("crimes_formatés_2012_2021.csv", index=False, encoding="utf-8-sig")

# Affichage d'un échantillon
print("Données traitées avec succès :")
print(df_final.head())
