import os
import tempfile
import unittest
from unittest.mock import patch

from excel_export import ExcelExporter


class ExportYearFilenameTests(unittest.TestCase):
    def test_person_export_uses_selected_year_in_filename(self):
        data = {
            "cat_biblio": None,
            "ind_biblio": None,
            "domaines": None,
            "detail": None,
        }

        with tempfile.TemporaryDirectory() as export_dir:
            with (
                patch("excel_export.pd.ExcelWriter") as excel_writer,
                patch("excel_export.ExcelHelper.export_origine"),
            ):
                ExcelExporter(export_dir).export_par_personne(
                    {"Alice": data},
                    "2025",
                )

        filename = os.path.basename(excel_writer.call_args.args[0])
        self.assertEqual(filename, "Alice_2025.xlsx")

    def test_library_export_uses_selected_year_in_filename(self):
        data = {
            "cat": None,
            "cat_biblio": None,
            "ind": None,
            "ind_biblio": None,
            "domaines": None,
            "detail": None,
        }

        with tempfile.TemporaryDirectory() as export_dir:
            with (
                patch("excel_export.pd.ExcelWriter") as excel_writer,
                patch("excel_export.ExcelHelper.export_origine"),
            ):
                ExcelExporter(export_dir).export_par_bibliotheque(
                    {"CENT": data},
                    "2025",
                )

        filename = os.path.basename(excel_writer.call_args.args[0])
        self.assertEqual(filename, "CENT_2025.xlsx")


if __name__ == "__main__":
    unittest.main()
