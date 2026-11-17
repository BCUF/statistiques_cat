import unittest

import pandas as pd

from transformations_df import create_other_dfs


class CreateOtherDfsTests(unittest.TestCase):
    def test_sparse_statistics_include_missing_categories_and_months(self):
        df = pd.DataFrame([
            {
                "Personne": "abc",
                "Bibliothèque": "CENT",
                "Mois": "04",
                "Traitement 1": "C",
                "Traitement 2": "_",
                "Traitement 3": "F",
                "Traitement 4": "R",
            },
            {
                "Personne": "abc",
                "Bibliothèque": "CENT",
                "Mois": "04",
                "Traitement 1": "I",
                "Traitement 2": "_",
                "Traitement 3": "I",
                "Traitement 4": "_",
            },
        ])

        (
            _,
            _,
            _,
            cat_par_personne,
            cat_par_biblio,
            ind_par_personne,
            ind_par_biblio,
            cent,
        ) = create_other_dfs(df)

        for stats in (cat_par_personne, cat_par_biblio):
            self.assertEqual(stats.loc[0, "A"], 0)
            self.assertEqual(stats.loc[0, "C"], 0)
            self.assertEqual(stats.loc[0, "01"], 0)
            self.assertEqual(stats.loc[0, "F"], 1)
            self.assertEqual(stats.loc[0, "04"], 1)

        for stats in (ind_par_personne, ind_par_biblio):
            self.assertEqual(stats.loc[0, "01"], 0)
            self.assertEqual(stats.loc[0, "04"], 1)
            self.assertEqual(stats.loc[0, "I"], 1)

        self.assertEqual(cent.loc[0, "02"], 0)
        self.assertEqual(cent.loc[0, "04"], 1)
        self.assertEqual(cent.loc[1, "02"], 0)
        self.assertEqual(cent.loc[1, "04"], 1)

        mois = [f"{mois:02d}" for mois in range(1, 13)]
        for stats in (
            cat_par_personne,
            cat_par_biblio,
            ind_par_personne,
            ind_par_biblio,
            cent,
        ):
            self.assertTrue(set(mois).issubset(stats.columns))


if __name__ == "__main__":
    unittest.main()
