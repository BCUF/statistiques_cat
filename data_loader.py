# Ce fichier permet d'importer le fichier Excel contenant les données brutes, elles-mêmes exportées depuis Alma Analytics.

import os
import datetime as dt
import pandas as pd

def load_data(input_file):

    if input_file is None:
        raise ValueError("Le chemin du fichier Excel source n'a pas été fourni.")
    else:
        path = os.path.abspath(input_file)
        if not os.path.isfile(path):
            raise FileNotFoundError(
                f"Le fichier suivant est introuvable : {path}"
            )


    # On importe le fichier Excel:
    df = pd.read_excel(path, engine="openpyxl")

    # La colonne "993 - Local Param 14" est renommée "Champs 993 - all":
    df = df.rename(
        columns={
            "993 - Local Param 14": "Champs 993 - all",
            "990 - Local Param 13": "Champs 990 - all"
        }
    )

    return df
