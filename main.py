"""Generate Excel reports from the cataloging statistics source workbook."""

import argparse
import ctypes

from pipeline import run_pipeline
from excel_export import ExcelExporter
from config import EXPORT_FOLDER, YEAR, validate_year


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


def _exporter_resultats(
    choix,
    resultats,
    output_folder=None,
    output_folders=None,
    progress_callback=None,
    year=YEAR,
):
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
    output_folders = output_folders or {}
    total_outputs = 0
    if "valid_fields" in products:
        total_outputs += 1
    if "employees" in products:
        total_outputs += len(resultats["person_files"])
    if "libraries" in products:
        total_outputs += len(resultats["bibli_files"])
    if "errors" in products:
        total_outputs += 1

    completed_outputs = 0
    if progress_callback is not None:
        progress_callback(0, total_outputs, "Préparation des fichiers...")

    def report_completed_output(description):
        nonlocal completed_outputs
        completed_outputs += 1
        if progress_callback is not None:
            progress_callback(
                completed_outputs,
                total_outputs,
                description,
            )

    def get_exporter(product):
        product_folder = output_folders.get(product, output_folder)
        return (
            ExcelExporter(product_folder)
            if product_folder is not None
            else ExcelExporter()
        )

    if "valid_fields" in products:
        get_exporter("valid_fields").export_corrects(
            resultats["cat_personne"],
            resultats["cat_biblio"],
            resultats["ind_personne"],
            resultats["ind_biblio"],
            resultats["cent"],
            resultats["origine"],
            resultats["cat"],
            resultats["ind"],
            year
        )
        report_completed_output("Fichier des champs corrects créé")

    if "employees" in products:
        get_exporter("employees").export_par_personne(
            resultats["person_files"],
            year,
            on_file_complete=lambda person: report_completed_output(
                f"Fichier de {person} créé"
            ),
        )

    if "libraries" in products:
        get_exporter("libraries").export_par_bibliotheque(
            resultats["bibli_files"],
            year,
            on_file_complete=lambda library: report_completed_output(
                f"Fichier de {library} créé"
            ),
        )

    if "errors" in products:
        get_exporter("errors").export_errors(df_errors)
        report_completed_output("Fichier des erreurs créé")


def _generer_exports(
    choix,
    input_file=None,
    output_folder=None,
    output_folders=None,
    year=YEAR,
    progress_callback=None,
):
    if progress_callback is not None:
        progress_callback(0, 0, "Analyse des données...")
    resultats = run_pipeline(input_file, year=year)
    _exporter_resultats(
        choix,
        resultats,
        output_folder,
        output_folders,
        progress_callback,
        year=year,
    )


def _parse_year(value):
    try:
        return validate_year(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError(str(error)) from error


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
    for product in ("errors", "valid_fields", "employees", "libraries"):
        parser.add_argument(
            f"--{product.replace('_', '-')}-output-dir",
            help=(
                f"Output directory for {product.replace('_', ' ')} "
                "exports (overrides --output-dir)"
            ),
        )
    parser.add_argument(
        "-p",
        "--product",
        nargs="+",
        choices=PRODUCT_CHOICES,
        help="one or more exports to generate: errors, valid_fields, employees, libraries or all",
    )
    parser.add_argument(
        "-y",
        "--year",
        type=_parse_year,
        help=f"Year to process (defaults to {YEAR})",
    )
    args = parser.parse_args(argv)

    if args.gui:
        if any(
            value is not None
            for value in (
                args.input_file,
                args.output_dir,
                args.errors_output_dir,
                args.valid_fields_output_dir,
                args.employees_output_dir,
                args.libraries_output_dir,
                args.product,
                args.year,
            )
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

    output_folders = {
        product: getattr(args, f"{product}_output_dir")
        for product in ("errors", "valid_fields", "employees", "libraries")
        if getattr(args, f"{product}_output_dir") is not None
    }
    if any(
        value is not None
        for value in (
            args.input_file,
            args.output_dir,
            args.product,
            args.year,
            *output_folders.values(),
        )
    ):
        if args.input_file is None:
            parser.error("--input-file is required to generate exports")
        if any(not folder.strip() for folder in output_folders.values()):
            parser.error("Product-specific output directories cannot be empty")

        products = (
            PRODUCT_CHOICES["all"]
            if args.product is None
            else set().union(*(PRODUCT_CHOICES[product] for product in args.product))
        )
        _generer_exports(
            products,
            input_file=args.input_file,
            output_folder=args.output_dir,
            output_folders=output_folders,
            year=args.year or YEAR,
        )
        print()
        print("Asked files have been generated in the configured output folders.")
        return

    parser.print_help()


if __name__ == "__main__":
    main()
