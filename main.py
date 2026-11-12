"""Generate Excel reports from the cataloging statistics source workbook."""

import argparse
import ctypes

from pipeline import run_pipeline
from excel_export import ExcelExporter
from config import EXPORT_FOLDER, INPUT_FOLDER, MODE_EXPORT


EXPORT_CHOICES = {
    "1": "Tout",
    "2": "Fichier des erreurs uniquement",
    "3": "Fichier des champs corrects uniquement",
    "4": "Fichiers par personne uniquement",
    "5": "Fichiers par bibliothèque uniquement",
    "6": "Fichiers des personnes + bibliothèques",
    "7": "Fichiers des champs corrects + personnes + bibliothèques",
}


def demander_exports():
    print()
    print("Que souhaitez-vous générer ?")
    print()
    for choix, description in EXPORT_CHOICES.items():
        print(f"{choix} - {description}")
    print()

    return input("Votre choix : ").strip()


def _exporter_resultats(choix, resultats, output_folder=None):
    if choix not in EXPORT_CHOICES:
        raise ValueError(f"Choix non reconnu : {choix}")

    df_errors = (
        resultats["errors"]
        .astype(str)
        .set_index("Champs 993 - all")
    )
    exporter = (
        ExcelExporter(output_folder)
        if output_folder is not None
        else ExcelExporter()
    )

    if choix in {"1", "3", "7"}:
        exporter.export_corrects(
            resultats["cat_personne"],
            resultats["cat_biblio"],
            resultats["ind_personne"],
            resultats["ind_biblio"],
            resultats["cent"],
            resultats["origine"],
            resultats["cat"],
            resultats["ind"],
        )

    if choix in {"1", "4", "6", "7"}:
        exporter.export_par_personne(resultats["person_files"])

    if choix in {"1", "5", "6", "7"}:
        exporter.export_par_bibliotheque(resultats["bibli_files"])

    if choix in {"1", "2"}:
        exporter.export_errors(df_errors)


def _generer_exports(choix, input_folder=INPUT_FOLDER, output_folder=None):
    resultats = run_pipeline(input_folder)
    _exporter_resultats(choix, resultats, output_folder)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Génère les fichiers Excel de statistiques de catalogage."
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="lance l'interface graphique au lieu du menu texte",
    )
    args = parser.parse_args(argv)

    if args.gui:

        # Masque et détruit la fenêtre de console sous Windows
        ctypes.windll.kernel32.FreeConsole()

        from gui import lancer_interface

        lancer_interface(
            EXPORT_CHOICES,
            _generer_exports,
            INPUT_FOLDER,
            EXPORT_FOLDER,
        )
        return

    resultats = run_pipeline()
    choix = demander_exports() if MODE_EXPORT == "menu" else "1"

    try:
        _exporter_resultats(choix, resultats)
    except ValueError as error:
        print()
        print(error)
        print()
        return

    print()
    print("Les fichiers demandés ont été générés.")


if __name__ == "__main__":
    main()
