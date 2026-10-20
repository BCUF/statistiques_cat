# Ce fichier est nécessaire pour faire tourner les fichiers main.py, main_errors.py et main_corrects.py. Il permet en effet d'éviter d'avoir le même code dans les trois fichiers, en lançant run_pipeline(), et donne une vue d'ensemble du traitement complet qui est effectué.

import pandas as pd

from data_loader import load_data

from errors_detection import (
    detect_length_errors,
    detect_duplicates,
    detect_regex_errors,
    detect_future_dates
)

from transformations_df import (
    create_df_origine,
    create_other_dfs,
    create_person_files_data,
    create_bibli_files_data
)


def run_pipeline():
    """
    On exécute l'ensemble du traitement et on retourne toutes les données dont les différents main pourraient avoir besoin.
    """

    # Chargement des données brutes:
    df_orig = load_data()

    # Détection des erreurs liées à la longueur des champs 993:
    df_length_ok, err_length = detect_length_errors(df_orig)

    # Détection des doublons dans les champs 993:
    df_no_duplicates, err_dup = detect_duplicates(df_length_ok)

    # Détection des erreurs liées à l'expression régulière:
    df_regex_ok, err_regex = detect_regex_errors(
        df_no_duplicates,
        df_orig
    )

    # Détection des dates futuristes:
    err_dates = detect_future_dates(
        df_regex_ok,
        df_orig
    )

    # Création de la dataframe contenant les erreurs:
    df_errors = pd.concat(
        [
            err_length,
            err_dup,
            err_regex,
            err_dates
        ]
    )

    # Création de la dataframe "Origine":
    df_ok = create_df_origine(
        df_errors,
        df_orig
    )

    # Création des autres dataframes:
    (
        df_origine,
        df_cat,
        df_ind,
        df_final_cat_par_personne,
        df_final_cat_par_biblio,
        df_final_ind_par_personne,
        df_final_ind_par_biblio,
        df_final_cent
    ) = create_other_dfs(df_ok)

    # Création de person_files
    person_files = create_person_files_data(
        df_origine,
        df_final_cat_par_personne,
        df_final_ind_par_personne
    )

    # Création de biblio_files
    bibli_files = create_bibli_files_data(
        df_origine
    )

    # Création d'un dictionnaire Python afin de récupérer facilement les dataframes dans les deux fichiers main_errors.py et main_corrects.py:
    return {
        "errors": df_errors,
        "origine": df_origine,
        "cat": df_cat,
        "ind": df_ind,
        "cat_personne": df_final_cat_par_personne,
        "cat_biblio": df_final_cat_par_biblio,
        "ind_personne": df_final_ind_par_personne,
        "ind_biblio": df_final_ind_par_biblio,
        "cent": df_final_cent,
        "person_files": person_files,
        "bibli_files": bibli_files
    }