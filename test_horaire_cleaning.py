#!/usr/bin/env python3
"""
Test de la fonction clean_horaire_name
"""

def clean_horaire_name(sheet_name):
    """Nettoie le nom de la feuille pour extraire le nom d'horaire de manière cohérente."""
    # Cette fonction doit être identique à celle utilisée dans fenetre_principale.py et Assignation des Niveaux.py
    sheet_lower = sheet_name.lower()
    type_intervenant = "animateur" if "animateur" in sheet_lower else "professeur"

    # Nettoyer le nom en supprimant les mots-clés d'intervenants
    horaire = (sheet_name
               .replace("animateur", "")
               .replace("Animateur", "")
               .replace("professeur", "")
               .replace("Professeur", "")
               .strip())

    # Si le résultat est vide, utiliser le nom original
    return horaire or sheet_name

if __name__ == "__main__":
    test_cases = [
        '8h15 à 10h15 Professeur',
        '10h30 à 11h30 Animateur',
        '9h à 12h20',
        '13h30 à 16h Professeur quelque chose',
        'Professeur 8h15-9h15',
        '8h15-9h15 Animateur',
        '9h à 12h20 Professeur',
        '13h30 à 16h Animateur Test'
    ]

    print('Test de la fonction clean_horaire_name:')
    print('=' * 50)

    for test in test_cases:
        result = clean_horaire_name(test)
        print(f'Input:  "{test}"')
        print(f'Output: "{result}"')
        print('-' * 30)
