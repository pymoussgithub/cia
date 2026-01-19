#!/usr/bin/env python3
"""
Test final pour vérifier que les horaires sont identiques
entre fenetre_principale.py et Assignation des Niveaux.py
"""
import os
import sys
import pandas as pd

# Importer les fonctions depuis les deux scripts
sys.path.append('.')

# Fonction de nettoyage unifiée
def clean_horaire_name(sheet_name):
    """Nettoie le nom de la feuille pour extraire le nom d'horaire de manière cohérente."""
    sheet_lower = sheet_name.lower()
    type_intervenant = "animateur" if "animateur" in sheet_lower else "professeur"

    # Liste complète des mots à supprimer (avec variations)
    words_to_remove = [
        "animateur", "Animateur", "anim", "Anim",
        "professeur", "Professeur", "prof", "Prof"
    ]

    # Nettoyer le nom en supprimant tous les mots-clés d'intervenants
    horaire = sheet_name
    for word in words_to_remove:
        horaire = horaire.replace(word, "")

    # Nettoyer les espaces multiples et supprimer les espaces au début/fin
    horaire = " ".join(horaire.split()).strip()

    # Si le résultat est vide, utiliser le nom original
    return horaire or sheet_name

def test_horaire_extraction():
    """Teste l'extraction des horaires depuis un fichier Excel."""
    week_folder = "semaine_1"
    test_file = "ecole_a.xlsx"
    filepath = os.path.join(week_folder, test_file)

    if not os.path.exists(filepath):
        print(f"Fichier {filepath} non trouvé")
        return

    print("=== TEST D'EXTRACTION DES HORAIRES ===")
    print(f"Fichier testé: {test_file}")
    print()

    try:
        xl = pd.ExcelFile(filepath)
        print("Horaires extraits:")

        for sheet_name in xl.sheet_names:
            cleaned_horaire = clean_horaire_name(sheet_name)
            print(f"  Feuille: '{sheet_name}'")
            print(f"  Horaire nettoyé: '{cleaned_horaire}'")
            print()

    except Exception as e:
        print(f"Erreur: {e}")

    print("=== RÉSULTAT ===")
    print("Si tous les horaires nettoyés sont cohérents,")
    print("alors le problème est résolu !")

if __name__ == "__main__":
    test_horaire_extraction()
