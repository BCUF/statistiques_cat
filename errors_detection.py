# Dans ce fichier, on définit toutes les fonctions qui permettent de détecter les erreurs dans les champs 993.

import re
import datetime as dt
import pandas as pd

# On importe l'année, le nombre maximal de champs 993 estimé pour une ressource, et les regex pour la date, le catalogage et l'indexation ; ceux-ci sont définis dans le fichier config.py:
from config import MAX_993_FIELDS, DATE_PATTERN, CATALOGAGE_PATTERN, INDEXATION_PATTERN


########################################################
# 1. Détection des erreurs liées à la taille des champs
########################################################
'''
On définit la fonction detect_length_errors(df), qui va détecter si le(s) champ(s) 993 d'un MMS Id donné est/sont de la bonne longueur. Cette fonction prend notre dataframe d'origine, df, comme argument.
La réflexion pour la formule 27*(i+1)+2*i est la suivante:
S'il y a un seul champ 993 correct, alors celui-ci fera une taille de 27 caractères (i = 0) ;
S'il y a deux champs 993 corrects, alors ils auront à eux deux une taille de 27*2 caractères + 2 caractères qui sont le point-virgule et l'espace séparant les deux champs (i = 1) ;
S'il y a trois champs 993 corrects, alors ils auront à eux deux une taille de 27*3 caractères + 4 caractères qui sont les 2 point-virgule et les 2 espaces séparant les trois champs (i = 2). Et ainsi de suite...
On en arrive finalement à la formule générale 27*(i+1)+2*i, où i peut prendre des valeurs entre 0 et l'infini. On a limité ici la valeur maximale de i à la valeur définie dans le fichier config.py (MAX_993_FIELDS), soit 50 au moment d'écrire ces lignes, car il y a peu de chances qu'un seul MMS ID ait davantage de champs 993...
'''
def detect_length_errors(df):
    # Cette fonction retourne les lignes dont la longueur du champ 993 est incorrecte.

    # Pour éviter certains warnings Pandas (e.g. SettingWithCopyWarning):
    df = df.copy()

    # On enregistre les longueurs de champs 993 "correctes" dans valid_lengths:
    valid_lengths = [27*(i+1) + 2*i for i in range(MAX_993_FIELDS)]

    # Création de la colonne "Field length", qui contient la longueur du/des champ(s) 993 pour un MMS ID donné:
    df["Field length"] = df["Champs 993 - all"].str.len()

    # Création de la colonne "Valid length", qui détermine si la longueur des champs 993 calculée dans "Field length" est correcte:
    df["Valid length"] = df["Field length"].isin(valid_lengths)

    # Création de la dataframe df_wrong, qui ne contient que les lignes pour lesquelles la longueur du/des champ/s 993 est/sont incorrecte/s, et création dans cette dataframe de la colonne "Type d'erreur" qui, comme son nom l'indique, renseigne sur le type d'erreur observé:
    df_wrong = df[~df["Valid length"]].copy()
    df_wrong["Type d'erreur"] = "Champ(s) de longueur(s) incorrecte(s)"

    # Création de la dataframe df_ok, qui ne contient que les lignes pour lesquelles la longueur du/des champ/s 993 est/sont correcte/s, et élimination des colonnes qui ne sont plus nécessaires ("Field length" et "Valid length"):
    df_ok = df[df["Valid length"]].drop(columns=["Field length", "Valid length"])

    # La fonction retourne les dataframes df_ok et df_wrong, et pour df_wrong, on ne garde que les colonnes "Champs 993 - all", "MMS Id" et "Type d'erreur":
    return df_ok, df_wrong[["Champs 993 - all", "MMS Id", "Type d'erreur"]]


########################################################
# 2. Détection des doublons
########################################################
'''
On définit la fonction detect_duplicates(df), qui est chargée de détecter les éventuels champs 993 à double pour un même MMS ID. L'idée est un peu tordue mais fonctionne malgré tout: pour chaque MMS Id, on va compter le nombre de points-virgules qui séparent chaque champ 993. Soit n le nombre de points-virgules pour un MMS Id donné. Si les champs 993 sont uniques (i.e. pas de doublons), alors il devrait y avoir n+1 champs 993 pour cet MMS Id ; autrement dit, il y aura moins de points-virgules que de champs 993 uniques. S'il y a des doublons, alors il y aura davantage ou autant de points-virgules que de champs 993 uniques.
'''
def detect_duplicates(df):
    # Pour éviter certains warnings Pandas (e.g. SettingWithCopyWarning), et df2 sera utilisé plus bas dans cette même fonction:
    df = df.copy()
    df2 = df.copy()

    # Création de la colonne "Nb semi-colons", qui indique combien de points-virgules se trouvent dans la colonne "Champs 993 - all":
    df["Nb semi-colons"] = df["Champs 993 - all"].str.count(";")

    # Séparation des champs au niveau du point-virgule:
    df_split = df["Champs 993 - all"].str.split(";", expand=True)
    df = pd.concat([df[["MMS Id", "Nb semi-colons"]], df_split], axis=1)

    # Pivot (les colonnes deviennent lignes) et élimination des lignes pour lesquelles le champ 993 est vide (dropna). Ces "champs vides" proviennent du pivot, pour les cas où il y avait moins de champs 993 que pour d'autres MMS IDs. On va également supprimer les éventuels espaces qui se trouvent avant ou après les champs 993 et dont l'origine est inconnue (strip):
    df = df.melt(id_vars=["MMS Id", "Nb semi-colons"], value_name="Champ 993")
    df = df.dropna(subset=["Champ 993"])
    df["Champ 993"] = df["Champ 993"].str.strip()

    # Suppression des doublons (on verra ci-dessous à quoi cela sert)
    df_unique = (
        df
        .drop_duplicates(subset=["MMS Id", "Champ 993"])
        .copy()
    )

    # Test doublons: pour chaque MMS ID, on compare le nombre de fois que celui-ci apparaît avec le nombre de points-virgules comptabilisés dans le champ 993 d'origine correspondant. La dataframe df_dup contient tous les cas où des doublons sont recensés:
    df_unique["Count"] = df_unique.groupby("MMS Id")["Champ 993"].transform("count")
    df_unique["Unique"] = df_unique["Count"] > df_unique["Nb semi-colons"]

    df_dup_ids = df_unique[~df_unique["Unique"]]["MMS Id"].unique()
    df_dup = df2[df2["MMS Id"].isin(df_dup_ids)].copy()

    # Création de la colonne "Type d'erreur" qui, comme son nom l'indique, renseigne sur le type d'erreur observé:
    df_dup["Type d'erreur"] = "Champ à double"

    # La dataframe df_ok contient uniquement les lignes de la dataframe df qui n'ont pas de doublons:
    df_ok = df_unique[df_unique["Unique"]]
    return df_ok, df_dup[["Champs 993 - all", "MMS Id", "Type d'erreur"]]


########################################################
# 3. Détection des erreurs regex
########################################################
'''
On définit la fonction detect_regex_errors(df_ok, df_orig), qui prendra comme arguments la dataframe df_ok créée par detect_duplicates() et notre dataframe d'origine. Cette fonction détecte d'autres erreurs non détectées jusqu'à présent, notamment des éventuels mélanges entre le catalogage et l'indexation lorsqu'il s'agit d'indiquer quel traitement a été effectué (un exemple d'une telle erreur serait l'indication du traitement I_S_, car la lettre S ne peut pas être utilisée pour indiquer le mode d'indexation). Il s'agit donc de comparer les champs 993 à deux expressions régulières différentes: une pour le catalogage et une pour l'indexation.
'''
def detect_regex_errors(df_ok, df_orig):

    # Pour éviter certains warnings Pandas (e.g. SettingWithCopyWarning):
    df_ok = df_ok.copy()

    # Détecte les champs 993 qui ne correspondent pas aux expressions régulières attendues pour le catalogage et l'indexation. Ces dernières sont définies dans le fichier config.py et il ne faut donc rien modifier ici:
    pattern = re.compile(
        f"{CATALOGAGE_PATTERN}|{INDEXATION_PATTERN}"
    )

    # Création de la dataframe df_good, qui contient tous les cas corrects, ainsi que la dataframe df_bad, qui contient tous les cas pour lesquels une erreur a été repérée au niveau de l'une des expressions régulières testées ci-dessus:

    df_ok["Regex ok"] = df_ok["Champ 993"].str.match(pattern)
    mask = df_ok["Champ 993"].str.match(pattern)

    # On crée df_good qui contient les cas corrects, sans erreurs au niveau des expressions régulières:
    df_good = df_ok[mask]

    # On crée df_bad, qui contient les erreurs au niveau des expressions régulières:
    df_bad = (
        df_ok.loc[~mask, ["MMS Id"]]
        .merge(df_orig, on="MMS Id")
    )

    # Création de la colonne "Type d'erreur" qui, comme son nom l'indique, renseigne sur le type d'erreur observé:
    df_bad["Type d'erreur"] = "Erreur regex"

    return df_good, df_bad[["Champs 993 - all", "MMS Id", "Type d'erreur"]]



########################################################
# 4. Détection des dates futuristes
########################################################
'''
On définit la fonction detect_future_dates(df_ok, df_orig), qui prendra comme arguments la dataframe df_good créée par detect_regex_errors() et notre dataframe d'origine. Cette fonction détecte les cas où la date entrée dans un champ 993 est postérieure à la date du jour et retourne uniquement la dataframe df_bad contenant ces cas.
'''
def detect_future_dates(df_ok, df_orig):

    # Pour éviter certains warnings Pandas (e.g. SettingWithCopyWarning)
    df_ok = df_ok.copy()

    # On crée la colonne "Date brute" qui contient la date extraite du champ 993:
    df_ok["Date brute"] = df_ok["Champ 993"].str.extract(rf'({DATE_PATTERN})')

    # Conversion du format "string" de la date en format "date":
    df_ok["Date brute"] = pd.to_datetime(df_ok["Date brute"], format="%Y%m%d")

    # Récupération de la date du jour:
    present = dt.datetime.now()

    # On récupère les cas où la date du champ 993 n'est pas encore advenue:
    df_bad = df_ok[df_ok["Date brute"] >= present][["MMS Id"]]

    # On crée la dataframe qu'il nous faut vraiment pour les erreurs, avec une colonne "Type d'erreur" qui indique qu'il s'agit de dates futuristes:
    df_bad = df_bad.merge(df_orig, on="MMS Id")
    df_bad["Type d'erreur"] = "Date futuriste"

    return df_bad[["Champs 993 - all", "MMS Id", "Type d'erreur"]]