
'''Ce fichier, lorsqu'il est exécuté, génère deux fichiers Excel: le premier contient les erreurs repérées et le deuxième contient les feuilles Excel qui nous intéressent avec les champs "corrects" pour le catalogage et l'indexation. 

Si l'on souhaite avoir uniquement le fichier des erreurs, on pourra exécuter main_errors.py à la place de main.py.

Si l'on souhaite avoir uniquement le fichier contenant les feuilles Excel qui contiennent les champs "corrects", on pourra exécuter main_corrects.py à la place de main.py.
'''

# On importe les différentes fonctions définies dans nos différents fichiers:
from pipeline import run_pipeline

from excel_export import (
    export_errors,
    export_corrects,
    export_par_personne,
    export_par_bibliotheque
)

from config import MODE_EXPORT

# On définit demander_exports() pour afficher un menu au lancement du programme, afin de pouvoir choisir les fichiers qu'on souhaite générer:

def demander_exports():

    print()
    print("Que souhaitez-vous générer ?")
    print()
    print("1 - Tout")
    print("2 - Fichier des erreurs uniquement")
    print("3 - Fichier des champs corrects uniquement")
    print("4 - Fichiers par personne uniquement")
    print("5 - Fichiers par bibliothèque uniquement")
    print("6 - Fichiers des personnes + bibliothèques")
    print("7 - Fichiers des champs corrects + personnes + bibliothèques")
    print()

    choix = input("Votre choix : ").strip()

    return choix


# On définit le main(), soit la partie principale du programme:
def main():

    resultats = run_pipeline()

    # Préparation du fichier contenant les erreurs détectées. La dataframe est convertie au format texte (string) afin d'éviter qu'Excel ne convertisse les données (et notamment les MMS IDs) en utilisant l'écriture exponentielle:
    df_errors = (
        resultats["errors"]
        .astype(str)
        .set_index("Champs 993 - all")
    )

    # La valeur de MODE_EXPORT est définie dans config.py
    if MODE_EXPORT == "menu":
        choix = demander_exports()
    else:
        choix = "1"

        # ==================================
    # 1 - Tout
    # ==================================

    if choix == "1":

        export_corrects(
            resultats["cat_personne"],
            resultats["cat_biblio"],
            resultats["ind_personne"],
            resultats["ind_biblio"],
            resultats["cent"],
            resultats["origine"],
            resultats["cat"],
            resultats["ind"]
        )

        export_par_personne(
            resultats["person_files"]
        )

        export_par_bibliotheque(
            resultats["bibli_files"]
        )

        export_errors(
            df_errors
        )

    # ==================================
    # 2 - Erreurs uniquement
    # ==================================

    elif choix == "2":

        export_errors(
            df_errors
        )

    # ==================================
    # 3 - Champs corrects uniquement
    # ==================================

    elif choix == "3":

        export_corrects(
            resultats["cat_personne"],
            resultats["cat_biblio"],
            resultats["ind_personne"],
            resultats["ind_biblio"],
            resultats["cent"],
            resultats["origine"],
            resultats["cat"],
            resultats["ind"]
        )

    # ==================================
    # 4 - Personnes uniquement
    # ==================================

    elif choix == "4":

        export_par_personne(
            resultats["person_files"]
        )

    # ==================================
    # 5 - Bibliothèques uniquement
    # ==================================

    elif choix == "5":

        export_par_bibliotheque(
            resultats["bibli_files"]
        )

    # ==================================
    # 6 - Personnes + bibliothèques
    # ==================================

    elif choix == "6":

        export_par_personne(
            resultats["person_files"]
        )

        export_par_bibliotheque(
            resultats["bibli_files"]
        )

    # ==================================
    # 7 - Champs corrects + personnes + bibliothèques
    # ==================================

    elif choix == "7":

        export_corrects(
            resultats["cat_personne"],
            resultats["cat_biblio"],
            resultats["ind_personne"],
            resultats["ind_biblio"],
            resultats["cent"],
            resultats["origine"],
            resultats["cat"],
            resultats["ind"]
        )

        export_par_personne(
            resultats["person_files"]
        )

        export_par_bibliotheque(
            resultats["bibli_files"]
        )

    else:

        print()
        print("Choix non reconnu.")
        print()

    print()
    print("Les fichiers demandés ont été générés.")
    input("Appuyez sur Enter pour fermer cette fenêtre.")

# Le "if statement" suivant peut sembler mystérieux, mais il s'agit d'une bonne pratique dans le cas où main.py devrait être importé dans un autre fichier. Cela éviterait qu'il soit exécuté automatiquement dès son importation:
if __name__ == "__main__":
    main()
