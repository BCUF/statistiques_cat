# Ce fichier permet d'importer le fichier Excel contenant les données brutes, elles-mêmes exportées depuis Alma Analytics.

import os
import datetime as dt
import pandas as pd

# On importe depuis config.py les noms du dossier (INPUT_FOLDER) dans lequel se trouve le fichier contenant les données brutes, exportées depuis Alma Analytics (INPUT_FILENAME). Si ces noms doivent être modifiés, il faut donc les modifier dans config.py et pas ici!
from config import (
    INPUT_FILENAME,
    INPUT_FOLDER
)

def load_data():

    # On définit current_dir, i.e. le dossier courant -- celui dans lequel se trouve le programme Python, et donc ce fichier data_loader.py -- ainsi que le dossier parent:
    current_dir = os.path.dirname(__file__)
    parent_dir = os.path.dirname(current_dir)

    # On définit data_dir, i.e. le dossier dans lequel se trouvent les données, en utilisant le nom de dossier (INPUT_FOLDER) défini dans config.py:
    data_dir = os.path.join(
        parent_dir,
        INPUT_FOLDER
    )

    # Dans config.py, le nom du fichier importé comporte l'élément "{date}" qui est ici remplacé par la date du jour:
    fname = INPUT_FILENAME.format(
        date=f"{dt.datetime.today():%Y%m%d}"
    )

    # On définit path en joinant data_dir et le nom du fichier:
    path = os.path.join(data_dir, fname)

    # Si le fichier est mal nommé ou s'il existe un autre problème par rapport au chemin d'accès, un message d'erreur apparaît au moment de lancer le programme:
    if not os.path.exists(path):
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
