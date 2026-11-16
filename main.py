"""Generate Excel reports from the cataloging statistics source workbook."""

import argparse
import ctypes

from pipeline import run_pipeline
from excel_export import ExcelExporter
from config import EXPORT_FOLDER


EXPORT_CHOICES = {
    "1": "Tout",
    "2": "Fichier des erreurs uniquement",
    "3": "Fichier des champs corrects uniquement",
    "4": "Fichiers par personne uniquement",
    "5": "Fichiers par bibliothèque uniquement",
    "6": "Fichiers des personnes + bibliothèques",
    "7": "Fichiers des champs corrects + personnes + bibliothèques",
}

PRODUCT_CHOICES = {
    "errors": {"errors"},
    "valid_fields": {"valid_fields"},
    "employees": {"employees"},
    "libraries": {"libraries"},
    "all": {"errors", "valid_fields", "employees", "libraries"},
}

EXPORT_CHOICE_PRODUCTS = {
    "1": PRODUCT_CHOICES["all"],
    "2": PRODUCT_CHOICES["errors"],
    "3": PRODUCT_CHOICES["valid_fields"],
    "4": PRODUCT_CHOICES["employees"],
    "5": PRODUCT_CHOICES["libraries"],
    "6": {"employees", "libraries"},
    "7": {"valid_fields", "employees", "libraries"},
}


def _exporter_resultats(choix, resultats, output_folder=None):
    if isinstance(choix, str):
        if choix not in EXPORT_CHOICES:
            raise ValueError(f"Choix non reconnu : {choix}")
        products = EXPORT_CHOICE_PRODUCTS[choix]
    else:
        products = set(choix)
        unknown_products = products.difference(
            {"errors", "valid_fields", "employees", "libraries"}
        )
        if unknown_products:
            raise ValueError(
                f"Produits non reconnus : {', '.join(sorted(unknown_products))}"
            )

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

    if "valid_fields" in products:
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

    if "employees" in products:
        exporter.export_par_personne(resultats["person_files"])

    if "libraries" in products:
        exporter.export_par_bibliotheque(resultats["bibli_files"])

    if "errors" in products:
        exporter.export_errors(df_errors)


def _generer_exports(
    choix,
    input_file=None,
    output_folder=None
):
    resultats = run_pipeline(input_file)
    _exporter_resultats(choix, resultats, output_folder)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Tools to generate Excel reports from the cataloging statistics source workbook."
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="lauch gui interface (cannot be combined with other options)",
    )
    parser.add_argument(
        "-i",
        "--input-file",
        help="Complete path to the source Excel file (exported from Alma Analytics)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        help=f"Output directory for the generated Excel files",
    )
    parser.add_argument(
        "-p",
        "--product",
        nargs="+",
        choices=PRODUCT_CHOICES,
        help="one or more exports to generate: errors, valid_fields, employees, libraries or all",
    )
    args = parser.parse_args(argv)

    if args.gui:
        if any(
            value is not None
            for value in (args.input_file, args.output_dir, args.product)
        ):
            parser.error("--gui cannot be combined with other options")

        # Masque et détruit la fenêtre de console sous Windows
        #ctypes.windll.kernel32.FreeConsole()

        from gui import lancer_interface

        lancer_interface(
            EXPORT_CHOICES,
            _generer_exports,
            "",
            EXPORT_FOLDER,
        )
        return

    if any(
        value is not None
        for value in (args.input_file, args.output_dir, args.product)
    ):
        if args.input_file is None:
            parser.error("--input-file is required to generate exports")

        products = (
            PRODUCT_CHOICES["all"]
            if args.product is None
            else set().union(*(PRODUCT_CHOICES[product] for product in args.product))
        )
        _generer_exports(
            products,
            input_file=args.input_file,
            output_folder=args.output_dir
        )
        print()
        print("Asked files have been generated in the output folder.")
        return

    parser.print_help()


if __name__ == "__main__":
    main()
