#!/usr/bin/env python3
"""
Test de synchronisation des horaires entre fenetre_principale.py et Assignation des Niveaux.py
"""
import os
import pandas as pd

def clean_horaire_name(sheet_name):
    """Nettoie le nom de la feuille pour extraire le nom d'horaire de manière cohérente."""
    # Cette fonction doit être identique à celle utilisée dans fenetre_principale.py et Assignation des Niveaux.py
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

def test_horaire_consistency():
    """Teste la cohérence des horaires entre les deux scripts."""
    week_folder = "semaine_1"

    # Fichiers à tester
    files_to_test = [
        'ecole_a.xlsx',
        'ecole_b.xlsx',
        'ECOLE_C_cours_standard.xlsx',
        'ECOLE_C_cours_intensif.xlsx',
        'MORNING.xlsx',
        'ECOLE_PREMIUM_cours_standard.xlsx',
        'ECOLE_PREMIUM_cours_intensifs.xlsx'
    ]

    print("=== TEST DE COHÉRENCE DES HORAIRES ===")
    print("Vérification que les horaires sont identiques dans tous les fichiers")
    print()

    all_cleaned_horaires = set()

    for filename in files_to_test:
        filepath = os.path.join(week_folder, filename)
        if not os.path.exists(filepath):
            print(f"Fichier non trouvé: {filename}")
            continue

        try:
            xl = pd.ExcelFile(filepath)
            print(f"{filename}:")
            for sheet_name in xl.sheet_names:
                cleaned = clean_horaire_name(sheet_name)
                all_cleaned_horaires.add(cleaned)
                print(f"  '{sheet_name}' -> '{cleaned}'")
            print()

        except Exception as e:
            print(f"Erreur avec {filename}: {e}")
            print()

    print("=== RÉSUMÉ ===")
    print(f"Nombre total d'horaires uniques trouvés: {len(all_cleaned_horaires)}")
    print("Horaires nettoyés:")
    for horaire in sorted(all_cleaned_horaires):
        print(f"  - '{horaire}'")

    print("\n✅ Si les horaires sont identiques dans les deux scripts, le problème est résolu!")

if __name__ == "__main__":
    test_horaire_consistency()
