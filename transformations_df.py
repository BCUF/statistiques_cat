# Dans ce fichier, on définit les fonctions qui vont créer les dataframes dont on aura besoin dans notre programme, et qui permettront d'obtenir nos différentes feuilles Excel.

import pandas as pd
import re

# Importation de CAT_IND_PATTERN et TRAITEMENT_PATTERN définis dans config.py:
from config import (
    CAT_IND_PATTERN,
    TRAITEMENT_PATTERN,
    DATE_SPLIT_PATTERN,
    CDD_MAPPING
)

###########################################################
# 0. Extraction des domaines CDD présents dans un champ 990
###########################################################

def extract_domains(champ_990):
    """
    Extrait les domaines CDD présents dans un champ 990.
    Exemple:    $$c cdd570 --> Biologie
    """

    if pd.isna(champ_990):
        return ["Domaine inconnu"]

    codes = re.findall(
        r"\$\$c\s+cdd\s*(\d{3})",
        str(champ_990),
        flags=re.IGNORECASE
    )

    # Si aucun domaine n'est trouvé:
    if not codes:
        return ["Domaine inconnu"]

    domaines = []

    for code in codes:

        domaine = CDD_MAPPING.get(
            code,
            f"cdd{code}"
        )

        domaines.append(domaine)

    # Suppression des doublons éventuels. Exemple: cdd900 présent deux fois ==> un seul domaine conservé

    return sorted(set(domaines))

########################################################
# 1. Création de la dataframe dont on aura besoin pour
# séparer toutes les parties des champs 993
########################################################

def create_df_origine(df, df_orig):
    # Afin d'éviter d'éventuels warnings:
    df = df.copy()
    df_orig = df_orig.copy()

    # On filtre la dataframe d'origine en éliminant les MMS Ids qui contiennent des erreurs (ceux-ci se trouvent dans df)
    df_filtered = df_orig.merge(df[['MMS Id']], on='MMS Id', how='left', indicator=True).query('_merge == "left_only"').drop(columns=['_merge'])

    # En cas de problème au niveau du filtrage, on affiche une erreur claire:
    if df_filtered.empty:
        raise ValueError(
            "Aucune notice valide après filtrage. "
            "Vérifier les regex et les fonctions de détection d'erreurs."
        )

    # On ne garde que les colonnes qui nous intéressent pour la suite
    df_filtered = df_filtered[
        [
            "Champs 993 - all",
            "MMS Id",
            "Champs 990 - all"
        ]
    ]

    # On transforme les champs 993 de façon à ce qu'on ait un seul champ par ligne
    df_filtered['Champs 993 - all'] = df_filtered['Champs 993 - all'].str.split(';')
    df_long = df_filtered.explode('Champs 993 - all')

    # On renomme la colonne "Champs 993 - all" en "Champ 993":
    df = df_long.rename(columns={"Champs 993 - all": "Champ 993", "Champs 990 - all": "Champ 990"})

    # Pour que les notices sans champ 990 soient traitées comme chaîne vide:
    df["Champ 990"] = (df["Champ 990"].fillna(""))

    # On découpe les champs 993 en ses différentes composantes pour pouvoir en extraire les informations. CAT_IND_PATTERN est défini dans config.py et correspond à l'expression régulière permettant de séparer les champs 993 dans leurs différentes composantes:
    df[["Bibliothèque", "Date brute", "Personne", "Traitement brut"]] = \
    df["Champ 993"].str.extract(CAT_IND_PATTERN)

    # Nettoyage du nom de la bibliothèque (on supprime les éventuels "_")
    df["Bibliothèque"] = df["Bibliothèque"].str.rstrip("_")

    # Création des colonnes pour les 4 traitements. TRAITEMENT_PATTERN est défini dans config.py, dans le cas où celui-ci aurait besoin d'être modifié à l'avenir:
    df[["Traitement 1", "Traitement 2", "Traitement 3", "Traitement 4"]] = \
    df["Traitement brut"].str.extract(TRAITEMENT_PATTERN)

    # Découpe de la date, en utilisant DATE_SPLIT_PATTERN, défini dans config.py:
    df[["Année", "Mois", "Jour"]] = (
        df["Date brute"]
        .str.extract(DATE_SPLIT_PATTERN)
    )

    # Création de la colonne "Date", qui contient une date lisible (au format JJ.MM.AAAA):
    df["Date"] = (
        df["Jour"].fillna("").astype(str)
        + "."
        + df["Mois"].fillna("").astype(str)
        + "."
        + df["Année"].fillna("").astype(str)
    )

    # Liste des domaines pour les calculs statistiques
    df["Domaines_liste"] = (
        df["Champ 990"]
        .apply(extract_domains)
    )

    # Version texte pour les exports Excel
    df["Domaines"] = (
        df["Domaines_liste"]
        .apply(lambda x: "; ".join(x))
    )


    # On retourne la dataframe qui contient les colonnes suivantes:
    return df[
        [
            "MMS Id",
            "Champ 993",
            "Bibliothèque",
            "Jour",
            "Mois",
            "Année",
            "Date",
            "Personne",
            "Traitement brut",
            "Traitement 1",
            "Traitement 2",
            "Traitement 3",
            "Traitement 4",
            "Champ 990",
            "Domaines"
        ]
    ]

########################################################
# 2. Création des autres dataframes dont on aura besoin
# pour le fichier Excel contenant les champs corrects
########################################################

def create_other_dfs(df):
    # Pour éviter des warnings:
    df = df.copy()

    # On crée les dataframes qui ne contiennent que les données d'origine pour le catalogage et l'indexation:
    df_cat = df[df["Traitement 1"].isin(['C', 'R', 'N'])]
    df_ind = df[df["Traitement 1"].isin(['I'])]

    # On calcule le nombre total de livres catalogués par personne:
    total_cat = df_cat.groupby('Personne').size().reset_index(name='Nombre total catalogage')

    # On calcule le nombre total de livres indexés par personne:
    total_ind = df_ind.groupby('Personne').size().reset_index(name='Nombre total indexation')

    # On calcule le nombre total de livres catalogués par biblio:
    total_cat_bib = df_cat.groupby('Bibliothèque').size().reset_index(name='Nombre total catalogage')

    # On calcule le nombre total de livres indexés par biblio:
    total_ind_bib = df_ind.groupby('Bibliothèque').size().reset_index(name='Nombre total indexation')

    # --------------------------------------------------
    # CATALOGAGE PAR PERSONNE: df_final_cat_par_personne
    # --------------------------------------------------

    # On crée une liste des 4 colonnes correspondant aux codes de traitement et au mode d'acquisition. Par exemple, si le traitement brut est "C_FR", on aura:
    # Traitement 1 = C
    # Traitement 2 = _
    # Traitement 3 = F
    # Traitement 4 = R
    colonnes_lettres = ['Traitement 1', 'Traitement 2', 'Traitement 3', 'Traitement 4']

    # On crée la liste "compteurs", vide pour l'instant, qui contiendra plusieurs dataframes intermédiaires. Chacune comptera les occurrences d'une partie du traitement:
    compteurs = []

    # Pour chacune des 4 colonnes du traitement, on compte combien de fois chaque valeur apparaît pour chaque personne. Plus précisément: pour chaque colonne de traitement, i.e. pour "Traitement 1", "Traitement 2", etc.:
    for col in colonnes_lettres:
        tmp = (
            # ...on regroupe les lignes par personne. P.ex. pour la personne abc, on aura C, C, R, etc.:
            df_cat.groupby('Personne')[col]

            # on compte le nombre d'occurrences de chaque lettre de traitement: par exemple:
            # abc C 2
            # abc R 1
            .value_counts()

            # On transforme les valeurs obtenues en colonnes, par exemple:
            #           C   N   R
            # Personne
            # abc       2   0   1
            .unstack(fill_value=0)

            # On fait en sorte que "Personne" redevienne une colonne normale, par exemple:
            # Personne  C   N   R
            # abc       2   0   1
            .reset_index()
        )
        
        # On ajoute le nom de la colonne d'origine devant chaque
        # valeur afin d'éviter que plusieurs colonnes portent le
        # même nom après les futures fusions. Par exemple:
        # C devient Traitement 1_C
        # N devient Traitement 1_N
        tmp = tmp.add_prefix(f"{col}_")

        # La colonne "Personne" a aussi reçu un préfixe, donc on lui redonne juste son nom initial:
        tmp = tmp.rename(columns={f"{col}_Personne": "Personne"})

        # On stocke cette dataframe intermédiaire dans la liste "compteurs". Comme on est dans une boucle "for" qui traite chaque traitement à tour de rôle, on aura 4 dataframes stockées dans compteurs:
        compteurs.append(tmp)

    # On calcule le nombre de livres catalogués par personne et par mois:
    par_mois = (
        df_cat.groupby(['Personne', 'Mois'])
        .size()
        .unstack(fill_value=0)
        .rename_axis(None, axis=1)
        .reset_index()
    )        

    # On initie la dataframe finale qui contiendra les totaux de catalogage par personne:
    df_final_cat_par_personne = total_cat

    # On ajoute les informations calculées précédemment dans la dataframe finale. Pour ce faire, on fait une boucle sur les 4 dataframes de traitement qu'on a stockées dans "compteurs" pour les rassembler en une seule dataframe:
    for c in compteurs:
        df_final_cat_par_personne = df_final_cat_par_personne.merge(c, on="Personne")

    # On rajoute les totaux par mois:
    df_final_cat_par_personne = df_final_cat_par_personne.merge(par_mois, on="Personne")

    # On remplace les noms de colonnes par des noms plus compréhensibles dans Excel:
    df_final_cat_par_personne = df_final_cat_par_personne.rename(columns={
        "Traitement 1_C": "Catalogage",
        "Traitement 1_N": "Notices sans inventaire",
        "Traitement 1_R": "Recatalogage",
        "Traitement 2_A": "Traitement pour une autre biblio",
        "Traitement 2__": "Rien à préciser",
        "Traitement 3_A": "A",
        "Traitement 3_D": "D",
        "Traitement 3_E": "E",
        "Traitement 3_F": "F",
        "Traitement 3_G": "G",
        "Traitement 3_L": "L",
        "Traitement 3_O": "O",
        "Traitement 3_P": "P",
        "Traitement 3_S": "S",
        "Traitement 3_Z": "Z",
        "Traitement 3__": "_",
        "Traitement 3_K": "K",
        "Traitement 4_C": "C",
        "Traitement 4_R": "R",
        "Traitement 4_I": "I",
        "Traitement 4_V": "V"
    })            


    # --------------------------------------------------
    # INDEXATION PAR PERSONNE: df_final_ind_par_personne
    # --------------------------------------------------

    # Afin d'alléger ce code, on se référera aux commentaires de la section précédente, intitulée "CATALOGAGE PAR PERSONNE: df_final_cat_par_personne", si l'on rencontre des problèmes pour comprendre ce qui se passe ci-dessous.

    # On remet à "zéro" la liste compteurs, déjà utilisée plus haut:
    compteurs = []

    # Pour chacune des 4 composantes du traitement de l'indexation, on compte le nombre d'occurrences de chaque valeur pour chaque personne.
    for col in colonnes_lettres:
        tmp = (
            df_ind.groupby('Personne')[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )

        # On ajoute le nom de la colonne d'origine devant chaque valeur afin d'éviter que plusieurs colonnes portent le même nom après les futures fusions. Par exemple:
        # I devient Traitement 1_I
        # N devient Traitement 3_N
        tmp = tmp.add_prefix(f"{col}_")

        # La colonne "Personne" a aussi reçu un préfixe, donc on lui redonne juste son nom initial:
        tmp = tmp.rename(columns={f"{col}_Personne": "Personne"})

        # On stocke cette dataframe intermédiaire dans la liste "compteurs":
        compteurs.append(tmp)

    # On calcule le nombre de ressources indexées par personne et par mois:
    par_mois = (
        df_ind.groupby(['Personne', 'Mois'])
        .size()
        .unstack(fill_value=0)
        .rename_axis(None, axis=1)
        .reset_index()
    )        

    # On initie la dataframe finale qui contiendra les totaux d'indexation par personne:
    df_final_ind_par_personne = total_ind

    # On ajoute les statistiques intermédiaires calculées dans les boucles précédentes:
    for c in compteurs:
        df_final_ind_par_personne = df_final_ind_par_personne.merge(c, on="Personne")

    # On rajoute les totaux par mois:
    df_final_ind_par_personne = df_final_ind_par_personne.merge(par_mois, on="Personne")

    '''
    Certaines colonnes ne sont pas utiles dans le résultat final. En effet, la colonne "Traitement 1_I" n'est pas utile car toutes les lignes d'indexation ont obligatoirement la valeur "I" dans cette colonne ; la colonne "Traitement 4__" n'est pas utile non plus car toutes les lignes d'indexation ont obligatoirement la valeur "_" dans cette colonne.
    errors="ignore" permet d'éviter une erreur si l'une de ces colonnes n'existe pas. Cela peut sembler inutile et contre-intuitif, mais ayant dû débugger ce code, je sais que cette ligne peut être utile et faire gagner du temps en cas de problème.
    '''
    df_final_ind_par_personne = df_final_ind_par_personne.drop(
        columns=["Traitement 1_I", "Traitement 4__"],
        errors="ignore"
    )

    # On remplace les noms de colonnes par des noms plus compréhensibles dans Excel:
    df_final_ind_par_personne = df_final_ind_par_personne.rename(columns={
        "Traitement 2_A": "Traitement pour une autre biblio",
        "Traitement 2__": "Rien à préciser",
        "Traitement 3_I": "I",
        "Traitement 3_J": "J",
        "Traitement 3_N": "N"
    })    


    # ----------------------------------------------
    # CATALOGAGE PAR BIBLIO: df_final_cat_par_biblio
    # ----------------------------------------------

    # Toujours afin d'alléger la suite de ce code, merci de se référer aux commentaires de la section "CATALOGAGE PAR PERSONNE: df_final_cat_par_personne" afin de comprendre ce qui se passe ci-dessous, si cela devait s'avérer nécessaire. On notera cependant qu'il ne s'agit plus ici de compter le nombre de ressources cataloguées par personne, mais par bibliothèque:

    compteurs = []

    for col in colonnes_lettres:
        tmp = (
            df_cat.groupby('Bibliothèque')[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )
        
        tmp = tmp.add_prefix(f"{col}_")
        tmp = tmp.rename(columns={f"{col}_Bibliothèque": "Bibliothèque"})
        compteurs.append(tmp)

    par_mois = (
        df_cat.groupby(['Bibliothèque', 'Mois'])
        .size()
        .unstack(fill_value=0)
        .rename_axis(None, axis=1)
        .reset_index()
    )        

    df_final_cat_par_biblio = total_cat_bib

    for c in compteurs:
        df_final_cat_par_biblio = df_final_cat_par_biblio.merge(c, on="Bibliothèque")

    df_final_cat_par_biblio = df_final_cat_par_biblio.merge(par_mois, on="Bibliothèque")

    df_final_cat_par_biblio = df_final_cat_par_biblio.rename(columns={
        "Traitement 1_C": "Catalogage",
        "Traitement 1_N": "Notices sans inventaire",
        "Traitement 1_R": "Recatalogage",
        "Traitement 2_A": "Traitement pour une autre biblio",
        "Traitement 2__": "Rien à préciser",
        "Traitement 3_A": "A",
        "Traitement 3_D": "D",
        "Traitement 3_E": "E",
        "Traitement 3_F": "F",
        "Traitement 3_G": "G",
        "Traitement 3_L": "L",
        "Traitement 3_O": "O",
        "Traitement 3_P": "P",
        "Traitement 3_S": "S",
        "Traitement 3_Z": "Z",
        "Traitement 3__": "_",
        "Traitement 3_K": "K",
        "Traitement 4_C": "C",
        "Traitement 4_R": "R",
        "Traitement 4_I": "I",
        "Traitement 4_V": "V"
    })       


    # ----------------------------------------------
    # INDEXATION PAR BIBLIO: df_final_ind_par_biblio
    # ----------------------------------------------

    # Toujours afin d'alléger la suite de ce code, merci de se référer aux commentaires de la section "CATALOGAGE PAR PERSONNE: df_final_cat_par_personne" afin de comprendre ce qui se passe ci-dessous, si cela devait s'avérer nécessaire. On notera cependant qu'il ne s'agit plus ici de compter le nombre de ressources indexées par personne, mais par bibliothèque:

    compteurs = []

    for col in colonnes_lettres:
        tmp = (
            df_ind.groupby('Bibliothèque')[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )

        tmp = tmp.add_prefix(f"{col}_")
        tmp = tmp.rename(columns={f"{col}_Bibliothèque": "Bibliothèque"})
        compteurs.append(tmp)

    par_mois = (
        df_ind.groupby(['Bibliothèque', 'Mois'])
        .size()
        .unstack(fill_value=0)
        .rename_axis(None, axis=1)
        .reset_index()
    )        

    df_final_ind_par_biblio = total_ind_bib

    for c in compteurs:
        df_final_ind_par_biblio = df_final_ind_par_biblio.merge(c, on="Bibliothèque")

    df_final_ind_par_biblio = df_final_ind_par_biblio.merge(par_mois, on="Bibliothèque")

    df_final_ind_par_biblio = df_final_ind_par_biblio.drop(
        columns=["Traitement 1_I", "Traitement 4__"],
        errors="ignore"
        )

    df_final_ind_par_biblio = df_final_ind_par_biblio.rename(columns={
        "Traitement 2_A": "Traitement pour une autre biblio",
        "Traitement 2__": "Rien à préciser",
        "Traitement 3_I": "I",
        "Traitement 3_J": "J",
        "Traitement 3_N": "N"
    })    

    # ------------------------------------------
    # CAT et IND pour la centrale: df_final_CENT
    # ------------------------------------------

    # On crée ici la dataframe qui permettra de créer la feuille "CENT", indiquant le nombre total de ressources cataloguées et indexées par la BCU-Centrale pour l'année en cours et pour chaque mois où des données ont été collectées.

    # On filtre la dataframe "df" pour que seules les données de la Centrale (CENT) y figurent:
    df_cent = df[df["Bibliothèque"] == "CENT"]

    # On récupère la liste des mois pour lesquels des données existent. La fonction "sorted()" garantit que ceux-ci seront classés dans l'ordre chronologique:
    mois_disponibles = sorted(df_cent['Mois'].unique())

    # On calcule le nombre total de ressources cataloguées et indexées. Si le traitement 1 = I, alors il s'agit d'indexation, et toute autre lettre signifierait qu'il s'agit de catalogage:
    total_cat_cent = (df_cent['Traitement 1'] != 'I').sum()
    total_ind_cent = (df_cent['Traitement 1'] == 'I').sum()

    # On compte le nombre de ressources cataloguées par mois.
    cat_cent_mois = (
        df_cent[df_cent['Traitement 1'] != 'I']
        .groupby('Mois')
        .size()
        .reindex(mois_disponibles, fill_value=0)
    )
    
    # On compte le nombre de ressources indexées par mois.
    ind_cent_mois = (
        df_cent[df_cent['Traitement 1'] == 'I']
        .groupby('Mois')
        .size()
        .reindex(mois_disponibles, fill_value=0)
    )  

    # Création de la structure de base pour la dataframe finale, qui contiendra deux lignes: une pour le catalogage et une pour l'indexation
    df_final_cent = pd.DataFrame({
        'CENT traitement': ['Catalogage', 'Indexation'],
        'Nombre total de livres traités': [total_cat_cent, total_ind_cent],
    })

    # Pour chaque mois existant, on ajoute une nouvelle colonne:
    for mois in mois_disponibles:
        df_final_cent[mois] = [
            cat_cent_mois[mois],
            ind_cent_mois[mois]
        ]

    # On retourne toutes les dataframes qui seront utilisées dans notre deuxième fichier Excel:
    return (
        df,
        df_cat,
        df_ind,
        df_final_cat_par_personne,
        df_final_cat_par_biblio,
        df_final_ind_par_personne,
        df_final_ind_par_biblio,
        df_final_cent
    )

def create_person_files_data(
        df_origine,
        df_final_cat_par_personne,
        df_final_ind_par_personne):
    """
    Prépare toutes les données nécessaires à la création
    d'un fichier Excel par personne.

    Pour chaque personne, on génère :

    - les données détaillées
    - les statistiques de catalogage
    - les statistiques d'indexation
    - les statistiques par domaine
    """

    result = {}

    personnes = sorted(
        df_origine["Personne"]
        .dropna()
        .unique()
    )

    for personne in personnes:

        # =====================================
        # DONNÉES DÉTAILLÉES
        # =====================================

        df_detail = (
            df_origine[
                df_origine["Personne"] == personne
            ]
            .copy()
        )

        # =====================================
        # STATISTIQUES PAR BIBLIOTHÈQUE
        # =====================================

        # =====================================
        # STATISTIQUES CATALOGAGE PAR BIBLIO
        # =====================================

        cat_biblio = (
            create_cat_stats_par_bibliotheque(
                df_detail,
                "Bibliothèque"
            )
        )

        # =====================================
        # STATISTIQUES INDEXATION PAR BIBLIO
        # =====================================

        ind_biblio = (
            create_ind_stats_par_bibliotheque(
                df_detail,
                "Bibliothèque"
            )
        )

        # =====================================
        # STATISTIQUES CATALOGAGE
        # =====================================

        df_cat = (
            df_final_cat_par_personne[
                df_final_cat_par_personne["Personne"]
                == personne
            ]
            .copy()
        )

        if df_cat.empty:
            df_cat = None

        # =====================================
        # STATISTIQUES INDEXATION
        # =====================================

        df_ind = (
            df_final_ind_par_personne[
                df_final_ind_par_personne["Personne"]
                == personne
            ]
            .copy()
        )

        if df_ind.empty:
            df_ind = None

        # =====================================
        # STATISTIQUES PAR DOMAINE
        # =====================================

        df_domaines = df_detail.copy()

        # Recréation de la liste des domaines à partir du champ 990
        df_domaines["Domaines_liste"] = (
            df_domaines["Champ 990"]
            .apply(extract_domains)
        )

        df_domaines = df_domaines.explode(
            "Domaines_liste"
        )

        # Réinitialisation de l'index pour éviter les doublons créés par explode()
        df_domaines = df_domaines.reset_index(drop=True)

        df_domaines["Type"] = (
            df_domaines["Traitement 1"]
            .map({
                "C": "Catalogage",
                "R": "Catalogage",
                "N": "Catalogage",
                "I": "Indexation"
            })
        )

        # Suppression des lignes incomplètes

        df_domaines = df_domaines.dropna(
            subset=["Domaines", "Type"]
        )

        stats_domaines = pd.crosstab(
            index=df_domaines["Domaines_liste"],
            columns=df_domaines["Type"]
        )

        # Création de la colonne Total seulement si qqn fait à la fois du catalogage et de l'indexation:
        if (
            "Catalogage" in stats_domaines.columns
            and
            "Indexation" in stats_domaines.columns
        ):

            stats_domaines["Total"] = (
                stats_domaines["Catalogage"]
                + stats_domaines["Indexation"]
            )


        # Transformation de l'index en colonne

        stats_domaines = (
            stats_domaines
            .reset_index()
        )

        # Renomme la colonne Domaine
        stats_domaines = stats_domaines.rename(
            columns={
                "Domaines_liste": "Domaines"
            }
        )

        # Tri décroissant sur la colonne la plus pertinente

        if "Total" in stats_domaines.columns:

            stats_domaines = (
                stats_domaines
                .sort_values(
                    by="Total",
                    ascending=False
                )
            )

        elif "Catalogage" in stats_domaines.columns:

            stats_domaines = (
                stats_domaines
                .sort_values(
                    by="Catalogage",
                    ascending=False
                )
            )

        elif "Indexation" in stats_domaines.columns:

            stats_domaines = (
                stats_domaines
                .sort_values(
                    by="Indexation",
                    ascending=False
                )
            )

        # =====================================
        # STOCKAGE DES RÉSULTATS
        # =====================================

        result[personne] = {

            "detail": df_detail,

            "cat": df_cat,

            "ind": df_ind,

            "cat_biblio": cat_biblio,

            "ind_biblio": ind_biblio,

            "domaines": stats_domaines
        }

    return result


##########################################################################
# Cas spécial COLLUNI: jusqu'en juin 2026, nous avions EOC, MUS, BLL et BLE qui cataloguaient et indexaient chacune de leur côté. Depuis juillet 2026, ils cataloguent pour CENT en indiquant dans le champ 990 $$c si ce qu'ils traient concerne EOC, MUS ou LLGR (langues et littératures...). L'idée est de créer un fichier Excel pour colluni qui répertorie tous les cas de 2026, indépendamment de la façon dont ils ont été enregistrés dans les notices bib.
##########################################################################

def est_notice_colluni(champ_990):
    """
    Détermine si une notice CENT concerne les bibliothèques de COLLUNI.
    """

    if pd.isna(champ_990):
        return False

    champ_990 = str(champ_990)

    return any(
        code in champ_990
        for code in [
            "$$c MUS",
            "$$c EOC",
            "$$c LLGR"
        ]
    )


def get_colluni_site(champ_990):

    if pd.isna(champ_990):
        return None

    champ_990 = str(champ_990)

    if "$$c MUS" in champ_990:
        return "MUS"

    if "$$c EOC" in champ_990:
        return "EOC"

    if "$$c LLGR" in champ_990:
        return "LLGR"

    return None



def create_bibli_files_data(
        df_origine):
    """
    Prépare toutes les données nécessaires à la création
    d'un fichier Excel par bibliothèque.

    Pour chaque bibliothèque, on génère :

    - les données détaillées
    - les statistiques de catalogage
    - les statistiques d'indexation
    - les statistiques par domaine
    """

    result = {}

    bibliotheques = sorted(
        df_origine["Bibliothèque"]
        .dropna()
        .unique()
    )

    # Ajoute une bibliothèque virtuelle COLLUNI
    bibliotheques.append("COLLUNI")

    for biblio in bibliotheques:
    

        # =====================================
        # DONNÉES DÉTAILLÉES
        # =====================================

        # On conserve uniquement les notices :
        # - de la bibliothèque concernée
        # - réalisées par les personnes rattachées
        #   à cette bibliothèque

        if biblio == "COLLUNI":

            anciennes_biblio = [
                "MUS",
                "EOC",
                "BLE",
                "BLL"
            ]

            # Notices avant la fusion de BLE, BLL, MUS et EOC

            df_avant = (
                df_origine[
                    df_origine["Bibliothèque"]
                    .isin(anciennes_biblio)
                ]
                .copy()
            )

            # Notices après la fusion, signées CENT, pour COLLUNI

            df_apres = (
                df_origine[
                    (
                        df_origine["Bibliothèque"]
                        == "CENT"
                    )
                    &
                    (
                        df_origine["Champ 990"]
                        .apply(est_notice_colluni)
                    )
                ]
                .copy()
            )

            df_detail = pd.concat(
                [
                    df_avant,
                    df_apres
                ],
                ignore_index=True
            )

        else:

            df_detail = (
                df_origine[
                    df_origine["Bibliothèque"] == biblio
                ]
                .copy()
            )

        # Si aucune donnée, passer à la bibliothèque suivante

        if df_detail.empty:
            continue

        if biblio == "COLLUNI":

            df_colluni_sites = (
                df_detail.copy()
            )

            mask = (
                df_colluni_sites["Bibliothèque"]
                == "CENT"
            )

            df_colluni_sites.loc[
                mask,
                "Bibliothèques_COLLUNI"
            ] = (
                df_colluni_sites.loc[
                    mask,
                    "Champ 990"
                ]
                .apply(get_colluni_site)
            )

            df_colluni_sites[
                "Bibliothèques_COLLUNI"
            ] = (
                df_colluni_sites[
                    "Bibliothèques_COLLUNI"
                ]
                .fillna(
                    df_colluni_sites[
                        "Bibliothèque"
                    ]
                )
            )

            # =====================================
            # COLLUNI - catalogage par bibliothèque
            # =====================================

            df_cat_biblio_colluni = (
                create_cat_stats_par_bibliotheque(
                    df_colluni_sites,
                    "Bibliothèques_COLLUNI"
                )
            )

            # =====================================
            # COLLUNI - indexation par bibliothèque
            # =====================================

            df_ind_biblio_colluni = (
                create_ind_stats_par_bibliotheque(
                    df_colluni_sites,
                    "Bibliothèques_COLLUNI"
                )
            )

        # =====================================
        # CATALOGAGE
        # =====================================

        df_cat = create_cat_stats_biblio(
            df_detail
        )

        # =====================================
        # INDEXATION
        # =====================================

        df_ind = create_ind_stats_biblio(
            df_detail
        )

        # =====================================
        # DOMAINES
        # =====================================

        df_domaines = df_detail.copy()

        df_domaines["Domaines_liste"] = (
            df_domaines["Champ 990"]
            .apply(extract_domains)
        )

        df_domaines = df_domaines.explode(
            "Domaines_liste"
        )

        df_domaines = df_domaines.reset_index(
            drop=True
        )

        df_domaines["Type"] = (
            df_domaines["Traitement 1"]
            .map({
                "C": "Catalogage",
                "R": "Catalogage",
                "N": "Catalogage",
                "I": "Indexation"
            })
        )

        df_domaines = df_domaines.dropna(
            subset=[
                "Domaines_liste",
                "Type"
            ]
        )

        stats_domaines = pd.crosstab(
            index=df_domaines["Domaines_liste"],
            columns=df_domaines["Type"]
        )

        # Total seulement si les deux colonnes existent

        if (
            "Catalogage" in stats_domaines.columns
            and
            "Indexation" in stats_domaines.columns
        ):

            stats_domaines["Total"] = (
                stats_domaines["Catalogage"]
                + stats_domaines["Indexation"]
            )

        stats_domaines = (
            stats_domaines
            .reset_index()
            .rename(
                columns={
                    "Domaines_liste": "Domaines"
                }
            )
        )

        # Tri décroissant

        if "Total" in stats_domaines.columns:

            stats_domaines = (
                stats_domaines
                .sort_values(
                    by="Total",
                    ascending=False
                )
            )

        elif "Catalogage" in stats_domaines.columns:

            stats_domaines = (
                stats_domaines
                .sort_values(
                    by="Catalogage",
                    ascending=False
                )
            )

        elif "Indexation" in stats_domaines.columns:

            stats_domaines = (
                stats_domaines
                .sort_values(
                    by="Indexation",
                    ascending=False
                )
            )

        # =====================================
        # RÉSULTAT FINAL
        # =====================================

        if biblio != "COLLUNI":

            df_cat_biblio_colluni = None
            df_ind_biblio_colluni = None


        result[biblio] = {

            "detail": df_detail,

            "cat": df_cat,

            "ind": df_ind,

            "domaines": stats_domaines,

            "cat_biblio": df_cat_biblio_colluni,

            "ind_biblio": df_ind_biblio_colluni
        }

    return result


#############################################
# Code pour gérer COLLUNI qui va me faire pleurer
#############################################

def create_cat_stats_par_bibliotheque(
        df,
        colonne_biblio):
    """
    Statistiques de catalogage par bibliothèque.
    """

    df = df.copy()

    df_cat = df[
        df["Traitement 1"].isin(
            ["C", "R", "N"]
        )
    ]

    if df_cat.empty:
        return None

    total_cat = (
        df_cat.groupby(colonne_biblio)
        .size()
        .reset_index(
            name="Nombre total catalogage"
        )
    )

    colonnes_lettres = [
        "Traitement 1",
        "Traitement 2",
        "Traitement 3",
        "Traitement 4"
    ]

    compteurs = []

    for col in colonnes_lettres:

        tmp = (
            df_cat
            .groupby(colonne_biblio)[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )

        tmp = tmp.add_prefix(
            f"{col}_"
        )

        tmp = tmp.rename(
            columns={
                f"{col}_{colonne_biblio}":
                colonne_biblio
            }
        )

        compteurs.append(tmp)

    par_mois = (
        df_cat
        .groupby(
            [colonne_biblio, "Mois"]
        )
        .size()
        .unstack(fill_value=0)
        .rename_axis(
            None,
            axis=1
        )
        .reset_index()
    )

    df_final = total_cat

    for c in compteurs:

        df_final = (
            df_final
            .merge(
                c,
                on=colonne_biblio
            )
        )

    df_final = (
        df_final
        .merge(
            par_mois,
            on=colonne_biblio
        )
    )

    df_final = df_final.rename(
        columns={
            "Traitement 1_C": "Catalogage",
            "Traitement 1_N": "Notices sans inventaire",
            "Traitement 1_R": "Recatalogage",
            "Traitement 2_A": "Traitement pour une autre biblio",
            "Traitement 2__": "Rien à préciser",
            "Traitement 3_A": "A",
            "Traitement 3_D": "D",
            "Traitement 3_E": "E",
            "Traitement 3_F": "F",
            "Traitement 3_G": "G",
            "Traitement 3_K": "K",
            "Traitement 3_L": "L",
            "Traitement 3_O": "O",
            "Traitement 3_P": "P",
            "Traitement 3_S": "S",
            "Traitement 3_Z": "Z",
            "Traitement 3__": "_",
            "Traitement 4_C": "C",
            "Traitement 4_I": "I",
            "Traitement 4_R": "R",
            "Traitement 4_V": "V"
        }
    )

    colonnes_obligatoires = [
        "Catalogage",
        "Notices sans inventaire",
        "Recatalogage",
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "A", "D", "E", "F", "G",
        "K", "L", "O", "P", "S",
        "Z", "_",
        "C", "I", "R", "V"
    ]

    for col in colonnes_obligatoires:

        if col not in df_final.columns:
            df_final[col] = 0

    for mois in [
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]:

        if mois not in df_final.columns:
            df_final[mois] = 0

    ordre = [
        colonne_biblio,
        "Nombre total catalogage",
        "Catalogage",
        "Notices sans inventaire",
        "Recatalogage",
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "A", "D", "E", "F", "G",
        "K", "L", "O", "P", "S",
        "Z", "_",
        "C", "I", "R", "V",
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]

    df_final = df_final[
        [col for col in ordre if col in df_final.columns]
    ]

    return df_final.reset_index(drop=True)


def create_ind_stats_par_bibliotheque(
        df,
        colonne_biblio):
    """
    Statistiques d'indexation par bibliothèque.
    """

    df = df.copy()

    df_ind = df[
        df["Traitement 1"] == "I"
    ]

    if df_ind.empty:
        return None

    total_ind = (
        df_ind.groupby(colonne_biblio)
        .size()
        .reset_index(
            name="Nombre total indexation"
        )
    )

    colonnes_lettres = [
        "Traitement 1",
        "Traitement 2",
        "Traitement 3",
        "Traitement 4"
    ]

    compteurs = []

    for col in colonnes_lettres:

        tmp = (
            df_ind
            .groupby(colonne_biblio)[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )

        tmp = tmp.add_prefix(
            f"{col}_"
        )

        tmp = tmp.rename(
            columns={
                f"{col}_{colonne_biblio}":
                colonne_biblio
            }
        )

        compteurs.append(tmp)

    par_mois = (
        df_ind
        .groupby(
            [colonne_biblio, "Mois"]
        )
        .size()
        .unstack(fill_value=0)
        .rename_axis(
            None,
            axis=1
        )
        .reset_index()
    )

    df_final = total_ind

    for c in compteurs:

        df_final = (
            df_final
            .merge(
                c,
                on=colonne_biblio
            )
        )

    df_final = (
        df_final
        .merge(
            par_mois,
            on=colonne_biblio
        )
    )

    df_final = df_final.drop(
        columns=[
            "Traitement 1_I",
            "Traitement 4__"
        ],
        errors="ignore"
    )

    df_final = df_final.rename(
        columns={
            "Traitement 2_A": "Traitement pour une autre biblio",
            "Traitement 2__": "Rien à préciser",
            "Traitement 3_I": "I",
            "Traitement 3_J": "J",
            "Traitement 3_N": "N"
        }
    )

    colonnes_obligatoires = [
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "I",
        "J",
        "N"
    ]

    for col in colonnes_obligatoires:

        if col not in df_final.columns:
            df_final[col] = 0

    for mois in [
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]:

        if mois not in df_final.columns:
            df_final[mois] = 0

    ordre = [
        colonne_biblio,
        "Nombre total indexation",
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "I",
        "J",
        "N",
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]

    df_final = df_final[
        [col for col in ordre if col in df_final.columns]
    ]

    return df_final.reset_index(drop=True)



#############################################
# Calcul de stats de catalogage par personne
# pour une bibliothèque donnée
#############################################

def create_cat_stats_biblio(df):
    """
    Calcule les statistiques de catalogage par personne
    pour une bibliothèque donnée.

    Le dataframe d'entrée doit déjà être filtré
    sur la bibliothèque concernée.
    """

    df = df.copy()

    df_cat = df[
        df["Traitement 1"].isin(
            ["C", "R", "N"]
        )
    ]

    if df_cat.empty:
        return None

    total_cat = (
        df_cat.groupby("Personne")
        .size()
        .reset_index(
            name="Nombre total catalogage"
        )
    )

    colonnes_lettres = [
        "Traitement 1",
        "Traitement 2",
        "Traitement 3",
        "Traitement 4"
    ]

    compteurs = []

    for col in colonnes_lettres:

        tmp = (
            df_cat
            .groupby("Personne")[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )

        tmp = tmp.add_prefix(
            f"{col}_"
        )

        tmp = tmp.rename(
            columns={
                f"{col}_Personne":
                "Personne"
            }
        )

        compteurs.append(tmp)

    par_mois = (
        df_cat
        .groupby(
            ["Personne", "Mois"]
        )
        .size()
        .unstack(fill_value=0)
        .rename_axis(
            None,
            axis=1
        )
        .reset_index()
    )

    df_final = total_cat

    for c in compteurs:

        df_final = (
            df_final
            .merge(
                c,
                on="Personne"
            )
        )

    df_final = (
        df_final
        .merge(
            par_mois,
            on="Personne"
        )
    )

    df_final = (
        df_final.rename(
            columns={
                "Traitement 1_C": "Catalogage",
                "Traitement 1_N": "Notices sans inventaire",
                "Traitement 1_R": "Recatalogage",
                "Traitement 2_A": "Traitement pour une autre biblio",
                "Traitement 2__": "Rien à préciser",
                "Traitement 3_A": "A",
                "Traitement 3_D": "D",
                "Traitement 3_E": "E",
                "Traitement 3_F": "F",
                "Traitement 3_G": "G",
                "Traitement 3_K": "K",
                "Traitement 3_L": "L",
                "Traitement 3_O": "O",
                "Traitement 3_P": "P",
                "Traitement 3_S": "S",
                "Traitement 3_Z": "Z",
                "Traitement 3__": "_",
                "Traitement 4_C": "C",
                "Traitement 4_I": "I",
                "Traitement 4_R": "R",
                "Traitement 4_V": "V"
            }
        )
    )

    # Colonnes attendues dans les exports Excel
    colonnes_obligatoires = [
        "Catalogage",
        "Notices sans inventaire",
        "Recatalogage",
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "A",
        "D",
        "E",
        "F",
        "G",
        "K",
        "L",
        "O",
        "P",
        "S",
        "Z",
        "_",
        "C",
        "I",
        "R",
        "V"
    ]

    for col in colonnes_obligatoires:
        if col not in df_final.columns:
            df_final[col] = 0


    for mois in [
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]:

        if mois not in df_final.columns:
            df_final[mois] = 0


    ordre = [
        "Personne",
        "Nombre total catalogage",
        "Catalogage",
        "Notices sans inventaire",
        "Recatalogage",
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "A",
        "D",
        "E",
        "F",
        "G",
        "K",
        "L",
        "O",
        "P",
        "S",
        "Z",
        "_",
        "C",
        "I",
        "R",
        "V",
        "01",
        "02",
        "03",
        "04",
        "05",
        "06",
        "07",
        "08",
        "09",
        "10",
        "11",
        "12"
    ]

    df_final = df_final[
        [col for col in ordre if col in df_final.columns]
    ]

    df_final = df_final.reset_index(drop=True)

    return df_final





##########################################################################
# Calcul des stats d'indexation par personne pour une bibliothèque donnée
##########################################################################

def create_ind_stats_biblio(df):
    """
    Calcule les statistiques d'indexation par personne
    pour une bibliothèque donnée.

    Le dataframe d'entrée doit déjà être filtré
    sur la bibliothèque concernée.
    """

    df = df.copy()

    df_ind = df[
        df["Traitement 1"] == "I"
    ]

    if df_ind.empty:
        return None

    total_ind = (
        df_ind.groupby("Personne")
        .size()
        .reset_index(
            name="Nombre total indexation"
        )
    )

    colonnes_lettres = [
        "Traitement 1",
        "Traitement 2",
        "Traitement 3",
        "Traitement 4"
    ]

    compteurs = []

    for col in colonnes_lettres:

        tmp = (
            df_ind
            .groupby("Personne")[col]
            .value_counts()
            .unstack(fill_value=0)
            .reset_index()
        )

        tmp = tmp.add_prefix(
            f"{col}_"
        )

        tmp = tmp.rename(
            columns={
                f"{col}_Personne":
                "Personne"
            }
        )

        compteurs.append(tmp)

    par_mois = (
        df_ind
        .groupby(
            ["Personne", "Mois"]
        )
        .size()
        .unstack(fill_value=0)
        .rename_axis(
            None,
            axis=1
        )
        .reset_index()
    )

    df_final = total_ind

    for c in compteurs:

        df_final = (
            df_final
            .merge(
                c,
                on="Personne"
            )
        )

    df_final = (
        df_final
        .merge(
            par_mois,
            on="Personne"
        )
    )

    df_final = df_final.drop(
        columns=[
            "Traitement 1_I",
            "Traitement 4__"
        ],
        errors="ignore"
    )

    df_final = (
        df_final.rename(
            columns={
                "Traitement 2_A": "Traitement pour une autre biblio",
                "Traitement 2__": "Rien à préciser",
                "Traitement 3_I": "I",
                "Traitement 3_J": "J",
                "Traitement 3_N": "N"
            }
        )
    )

    # Colonnes attendues dans les exports Excel

    colonnes_obligatoires = [
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "I",
        "J",
        "N"
    ]

    for col in colonnes_obligatoires:

        if col not in df_final.columns:

            df_final[col] = 0

    # Colonnes des mois

    for mois in [
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]:

        if mois not in df_final.columns:

            df_final[mois] = 0

    # Ordre des colonnes

    ordre = [
        "Personne",
        "Nombre total indexation",
        "Traitement pour une autre biblio",
        "Rien à préciser",
        "I",
        "J",
        "N",
        "01",
        "02",
        "03",
        "04",
        "05",
        "06",
        "07",
        "08",
        "09",
        "10",
        "11",
        "12"
    ]

    df_final = df_final[
        [col for col in ordre if col in df_final.columns]
    ]

    df_final = df_final.reset_index(
        drop=True
    )

    return df_final


