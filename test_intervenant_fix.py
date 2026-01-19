#!/usr/bin/env python3
"""
Script de test pour vérifier que les intervenants sont correctement récupérés
au niveau des classes individuelles plutôt qu'au niveau de l'horaire.
"""

import os
import sys
sys.path.append(os.path.dirname(__file__))

from fenetre_principale import analyze_school_classes

def test_intervenant_analysis():
    """Teste la récupération des intervenants au niveau des classes."""

    # Utiliser un dossier semaine existant pour le test
    week_folders = [f for f in os.listdir('.') if f.startswith('semaine_')]
    if not week_folders:
        print("ERREUR: Aucun dossier semaine trouve pour le test")
        return

    # Prendre le premier dossier semaine disponible
    week_folder = week_folders[0]
    print(f"Test avec le dossier: {week_folder}")

    # Analyser les données des écoles
    school_data = analyze_school_classes(week_folder)

    if not school_data:
        print("Aucune donnee d'ecole trouvee")
        return

    print("Analyse des ecoles reussie")
    print("Resume des donnees:")

    for school_key, horaires in school_data.items():
        print(f"\nEcole: {school_key}")

        for horaire_info in horaires:
            horaire = horaire_info.get('horaire', 'Inconnu')
            intervenant_horaire = horaire_info.get('intervenant', 'Non specifie')
            classes = horaire_info.get('classes', [])

            print(f"  Horaire: {horaire}")
            print(f"  Intervenant (horaire): {intervenant_horaire}")
            print(f"  Nombre de classes: {len(classes)}")

            for classe_info in classes:
                classe_nom = classe_info.get('nom_classe', 'Inconnue')
                intervenant_classe = classe_info.get('intervenant', 'Non specifie')
                nb_eleves = classe_info.get('nb_eleves', 0)

                print(f"    - Classe: {classe_nom}")
                print(f"      Intervenant (classe): {intervenant_classe}")
                print(f"      Nombre d'eleves: {nb_eleves}")

                # Verifier si l'intervenant de la classe est different de "Non specifie"
                if intervenant_classe != "Non specifie":
                    print("      [OK] Intervenant assigne a cette classe!")
                else:
                    print("      [ATTENTION] Aucun intervenant assigne a cette classe")

    print("\nTest termine!")

if __name__ == "__main__":
    test_intervenant_analysis()