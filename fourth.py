import pandas as pd
import numpy as np
import re
import sys
import os
from datetime import datetime

def est_adresse_ip_valide(ip):
    pattern = r"^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    return bool(re.match(pattern, str(ip)))


def lire_excel(nom_fichier):
    try:
        return pd.read_excel(nom_fichier, header=None)
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier Excel : {e}")
        return None


def trouver_cellule(df, valeur):
    for index, row in df.iterrows():
        for col, cell_value in enumerate(row):
            if str(cell_value).strip() == valeur:
                return index, col
    return None, None


def extraire_donnees(df, cellule):
    row, col = trouver_cellule(df, cellule)
    if row is None or col is None:
        print(f"Cellule '{cellule}' non trouvée dans le fichier Excel.")
        return []
    return df.iloc[row + 1:, col].dropna().tolist()


def ajuster_longueurs(*listes):
    max_len = max(len(liste) for liste in listes)
    return [liste + [''] * (max_len - len(liste)) for liste in listes]


def ecrire_csv(sources, services, destinations, actions, nom_fichier_sortie):
    try:
        headers = ["Source", "Source NAT", "Destination", "Destination NAT",
                   "Service", "Description", "User", "Application", "Action"]

        # Ajuster les longueurs des listes
        sources, services, destinations, actions = ajuster_longueurs(sources, services, destinations, actions)

        # Créer un dictionnaire avec les données
        data = {
            'Source': sources,
            'Service': services,
            'Destination': destinations,
            'Action': actions
        }

        # Créer le DataFrame en utilisant le dictionnaire
        df = pd.DataFrame(data)

        # Ajouter les colonnes manquantes
        for header in headers:
            if header not in df.columns:
                df[header] = ''

        # Réorganiser les colonnes selon l'ordre des en-têtes
        df = df[headers]

        df.to_csv(nom_fichier_sortie, index=False)
        print(f"Les données ont été écrites dans {nom_fichier_sortie} avec les en-têtes spécifiés")
    except Exception as e:
        print(f"Erreur lors de l'écriture du fichier CSV : {e}")


def excel_vers_csv(fichier_excel, fichier_csv=None):
    df = lire_excel(fichier_excel)
    if df is not None:
        sources = [ip for ip in extraire_donnees(df, "IP") if est_adresse_ip_valide(ip)]
        services = extraire_donnees(df, "Service/Port Destination")
        destinations = extraire_donnees(df, "Destination")
        actions = extraire_donnees(df, "Action")

        if fichier_csv is None:
            # Générer un nom de fichier automatiquement si non spécifié
            nom_base = os.path.splitext(os.path.basename(fichier_excel))[0]
            date_actuelle = datetime.now().strftime("%Y%m%d_%H%M%S")
            fichier_csv = f"{nom_base}_{date_actuelle}.csv"

        ecrire_csv(sources, services, destinations, actions, fichier_csv)
        return fichier_csv


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <fichier_excel> [fichier_csv_sortie]")
        sys.exit(1)

    fichier_excel = sys.argv[1]
    fichier_csv = sys.argv[2] if len(sys.argv) > 2 else None

    fichier_sortie = excel_vers_csv(fichier_excel, fichier_csv)
    print(f"Le fichier de sortie a été créé : {fichier_sortie}")