import pandas as pd

def generate_individual_school_files():
    # Liste précise des 5 fichiers à générer
    tasks = [
        {
            "filename": "ECOLE_C_cours_standard.xlsx",
            "sheet_name": "8h30 à 11h50 Prof",
            "horaire": "8h30 à 11h50"
        },
        {
            "filename": "ECOLE_C_cours_intensif.xlsx",
            "sheet_name": "11h50 à 12h35 Prof",
            "horaire": "11h50 à 12h35"
        },
        {
            "filename": "MORNING.xlsx",
            "sheet_name": "9h à 12h20 Prof",
            "horaire": "9h à 12h20"
        },
        {
            "filename": "ECOLE_PREMIUM_cours_standard.xlsx",
            "sheet_name": "9h à 12h20 Prof",
            "horaire": "9h à 12h20"
        },
        {
            "filename": "ECOLE_PREMIUM_cours_intensifs.xlsx",
            "sheet_name": "13h30 à 16h Prof",
            "horaire": "Mardi et jeudi 13h30 à 16h"
        }
    ]

    # Structure des colonnes de votre fichier original
    columns = ["Nom de la classe", "Niveau", "Intervenant", "Rôle", "Liste des élèves"]

    # Données vides ou exemples pour respecter la structure
    sample_data = []

    for task in tasks:
        # Création du DataFrame
        df = pd.DataFrame(sample_data, columns=columns)
        
        # Génération du fichier Excel
        with pd.ExcelWriter(task["filename"], engine="openpyxl") as writer:
            # On utilise le sheet_name défini (max 31 car.)
            df.to_excel(writer, sheet_name=task["sheet_name"][:31], index=False)
            
        print(f"Généré : {task['filename']} (Onglet: {task['sheet_name']})")

if __name__ == "__main__":
    generate_individual_school_files()