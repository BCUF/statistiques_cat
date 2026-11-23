import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook

from excel_export import ExcelExporter
from excel_helper import ExcelHelper


class ColluniExportSheetTests(unittest.TestCase):
    def test_cataloguing_and_indexing_person_and_library_sheets_are_separate(self):
        cat_personne = pd.DataFrame([{
            "Personne": "Alice",
            "Nombre total catalogage": 1,
            "Catalogage": 1,
            "Notices sans inventaire": 0,
            "Recatalogage": 0,
            "Traitement pour une autre biblio": 0,
            "Rien à préciser": 0,
            "A": 0,
            "A2": 0,
            "A3": 0,
            "A4": 0,
            "A5": 0,
            "A6": 0,
            "C": 1,
            "C2": 0,
            "C3": 0,
            "C4": 0,
            "01": 1,
            "02": 0,
            "03": 0,
            "04": 0,
            "05": 0,
            "06": 0,
        }])
        cat_biblio = cat_personne.rename(
            columns={"Personne": "Bibliothèques_COLLUNI"}
        )
        ind_personne = pd.DataFrame([{
            "Personne": "Alice",
            "Nombre total indexation": 1,
            "Traitement pour une autre biblio": 0,
            "Rien à préciser": 1,
            "I": 1,
            "I2": 0,
            "I3": 0,
            "01": 1,
            "02": 0,
            "03": 0,
            "04": 0,
            "05": 0,
            "06": 0,
            "07": 0,
        }])
        ind_biblio = ind_personne.rename(
            columns={"Personne": "Bibliothèques_COLLUNI"}
        )
        data = {
            "cat": cat_personne,
            "cat_biblio": cat_biblio,
            "ind": ind_personne,
            "ind_biblio": ind_biblio,
            "domaines": None,
            "detail": None,
        }

        with tempfile.TemporaryDirectory() as export_dir:
            with patch.object(
                ExcelHelper,
                "export_origine",
            ):
                ExcelExporter(export_dir).export_par_bibliotheque(
                    {"COLLUNI": data},
                    "2026",
                )

            workbook = load_workbook(
                Path(export_dir) / "COLLUNI_2026.xlsx",
                read_only=True,
            )
            try:
                self.assertEqual(
                    workbook.sheetnames,
                    [
                        "C-nbr par personne",
                        "C-nbr par biblio",
                        "I-nbr par personne",
                        "I-nbr par biblio",
                    ],
                )
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
