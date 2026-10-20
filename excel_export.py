# Ce fichier contient les fonctions qui permettent de formater nos différentes feuilles Excel comme nous le souhaitons, ainsi que les fonctions permettant d'exporter nos dataframes vers les feuilles Excel correspondantes.

import os
import datetime as dt
import pandas as pd

from openpyxl.styles import Font

# On importe depuis config.py les noms des fichiers qui contiendront les erreurs détectées (ERROR_FILENAME), les champs corrects (CORRECT_FILENAME), et le nom du dossier dans lequel ces deux fichiers seront exportés (EXPORT_FOLDER):
from config import (
    ERROR_FILENAME,
    CORRECT_FILENAME,
    EXPORT_FOLDER
)

######################################################
# 1. Définition des formats pour les feuilles Excel
######################################################

# On crée une fonction qu'on appellera pour formater nos différentes feuilles Excel.

def create_formats(workbook):
    return {

        # Format 'bold': pour que le texte soit en gras:
        "bold": workbook.add_format({
            'bold': True
        }),

        # Format 'rotate': pour que le texte de certaines cellules subissent une rotation de 90 degrés:
        "rotate": workbook.add_format({
            'rotation': 90,
            'bold': True,
            'align': 'center'    
        }),

        # Format 'wrap': pour que le texte de certaines cellules soit en gras et subisse un retour à la ligne dans Excel:
        "wrap": workbook.add_format({
            'text_wrap': True,
            'bold': True
        }),

        # Format 'bold_right': pour que les cellules des en-têtes soient en gras et aient une bordure à droite plus épaisse:
        "bold_right": workbook.add_format({
            'bold': True,
            'right': 2,
            'right_color': 'black'
        }),

        # Format 'rotate_right': pour que les cellules de certains en-têtes subissent une rotation de 90 degrés et aient une bordure à droite plus épaisse:
        "rotate_right": workbook.add_format({
            'rotation': 90,
            'bold': True,
            'align': 'center',
            'right': 2,
            'right_color': 'black'
        }),

        # Format 'wrap_right': texte en gras avec bordure à droite plus épaisse et retour à la ligne:
        "wrap_right": workbook.add_format({
            'text_wrap': True,
            'bold': True,
            'right': 2,
            'right_color': 'black'
        }),

        # Format 'data_right': cellules avec uniquement une bordure à droite plus épaisse, sans autres modifications:
        "data_right": workbook.add_format({
            'right': 2,
            'right_color': 'black'
        }),

        # Format 'data_left': cellules avec uniquement une bordure à gauche plus épaisse, sans autres modifications:
        "data_left": workbook.add_format({
            'left': 2,
            'left_color': 'black'
        }),

        # Format 'text': pour éviter que les colonnes contenant les MMS IDs soient écrites en notation scientifique par Excel. Le symbole "@" indique à Excel de traiter les cellules comme du texte:
        "text": workbook.add_format({
            "num_format": "@"
        })
    }

##################################################################
# 2. Fonction qui ajuste la largeur des colonnes à leur contenu
##################################################################

def autofit_columns(df, worksheet, rotated_columns, wrapped_columns):

    for i, col in enumerate(df.columns):


        # Largeur maximale du contenu:
        #content_width = df[col].astype(str).map(len).max()

        content_width = max(
            len(str(value))
            for value in df[col].fillna("")
        )

        # Colonnes où l'en-tête a subi une rotation: on ignore l'en-tête:
        if col in rotated_columns:
            header_width = 0
            # On adapte la largeur en fonction du contenu:
            final_width = content_width * 1.1 + 1
            # On ajoute un espace supplémentaire pour la flèche du filtre:
            final_width += 2
            worksheet.set_column(i, i, final_width)
            continue

        # Colonnes où on a un retour à la ligne (wrap): on ignore l'en-tête:
        elif col in wrapped_columns:
            header_width = 0
            # On choisit une largeur maximale raisonnable pour une colonne avec wrap:
            max_width = 30
            calculated = content_width * 1.1 + 5
            final_width = min(calculated, max_width)
            # Ajout d'un espace pour la flèche du filtre dans Excel:
            final_width += 2
            worksheet.set_column(i, i, final_width)
            continue

        # Colonnes "normales", sans retour à la ligne: l'en-tête est pris en compte:
        else:
            header_width = len(col)

        final_width = max(content_width, header_width) * 1.1 + 1

        # Ajout d'un espace pour la flèche du filtre dans Excel:
        final_width += 2

        worksheet.set_column(i, i, final_width)



###############################################################################
# 3. Formatage des feuilles de catalogage/indexation par personne et par biblio
###############################################################################

def formater_feuille_statistique(
        writer,                     # objet ExcelWriter
        df,                         # la dataframe à exporter 
        nom_feuille,                # nom de la feuille Excel
        zones_fusion,  # dico indiquant quelles colonnes ont un en-tête fusionné
        rotated_columns,
        wrapped_columns,
        cols_with_border_names):

    # On transforme la dataframe en Excel:
    df.to_excel(
        writer,
        sheet_name=nom_feuille,
        index=False,
        startrow=1 # on commence à la 2ème ligne dans Excel (startrow=1, car Python commence à compter à partir de 0)
    )

    workbook = writer.book # Fichier Excel complet
    worksheet = writer.sheets[nom_feuille] # Feuille Excel courante

    # Création des formats:
    formats = create_formats(workbook)

    format_total = workbook.add_format({
    "bold": True,
    "top": 2
    })

    # ------------------------------------------------------------------
    # LIGNE 0 : mise en forme des cellules fusionnées (sur la ligne 0)
    # ------------------------------------------------------------------

    # On parcourt le dico zones_fusion ; col_name = nom de colonne de la df, texte = texte à afficher dans la cellule fusionnée, span = nombre de colonnes supplémentaires à ajouter dans la fusion:
    for col_name, (texte, span) in zones_fusion.items():

        # On cherche l'indice de la colonne dans la dataframe:
        col = df.columns.get_loc(col_name)

        # Fusion de cellules sur la ligne 0:
        worksheet.merge_range(
            0,                      # première ligne = 0
            col,                    # première colonne = col
            0,                      # dernière ligne = 0
            col + span,             # dernière colonne = col + span
            texte,            # le texte à afficher dans les cellules fusionnées
            formats["data_left"]    # la bordure à gauche sera plus épaisse
        )

    # -------------------------------------------------------
    # LIGNE 1 : mise en forme des en-têtes (sur la ligne 1)
    # -------------------------------------------------------

    # On parcout les colonnes de la dataframe ; col_index = position de la colonne, col_name = nom de la colonne:
    for col_index, col_name in enumerate(df.columns):

        # Si le nom de la colonne est dans la liste "rotated_columns", on applique le format "rotate":
        if col_name in rotated_columns:
            fmt = formats["rotate"]

        # Si le nom de la colonne est dans la liste "wrapped_columns", on applique le format "wrap":
        elif col_name in wrapped_columns:
            fmt = formats["wrap"]

        # Dans tous les autres cas, le texte des en-têtes est seulement mis en gras:
        else:
            fmt = formats["bold"]

        # Ecriture de l'en-tête dans Excel, sur la ligne 1:
        worksheet.write(1, col_index, col_name, fmt)

    # -------------------------------
    # Colonnes avec bordure épaisse
    # -------------------------------

    # On crée une liste vide qui contiendra les indices des colonnes avec bordure épaisse:
    cols_with_border = []

    for nom_colonne in cols_with_border_names:

        # Comme on le verra plus bas, on a défini des tuples qui permettent de définir où doivent être placées les bordures plus épaisses dans Excel. On vérifie donc ici si les colonnes sont des tuples et si tel est le cas, on utilise que le premier élément du tuple.
        if isinstance(nom_colonne, tuple):
            cols_with_border.append(
                df.columns.get_loc(nom_colonne[0]) - 1
            )
        else:
            cols_with_border.append(
                df.columns.get_loc(nom_colonne)
            )

    # On récupère les indices des colonnes fusionnées:
    fusion_cols = [
        df.columns.get_loc(col)
        for col in zones_fusion
    ]

    # On réécrit les colonnes de façon à ce que certaines d'entre elles aient une bordure plus épaisse:
    for col in cols_with_border:

        # Récupération du nom de la colonne:
        header = df.columns[col]

        # Si la colonne est dans "rotated_columns", on lui applique le format "rotate_right":
        if header in rotated_columns:
            fmt_header = formats["rotate_right"]

        # Si la colonne est dans "wrapped_columns", on lui applique le format "wrap_right":
        elif header in wrapped_columns:
            fmt_header = formats["wrap_right"]

        # Dans les autres cas, on applique le format "bold_right":
        else:
            fmt_header = formats["bold_right"]

        # On réécrit l'en-tête sur la ligne 1 pour qu'il ait des bordures plus épaisses à certains endroits:
        worksheet.write(1, col, header, fmt_header)

        # On récupère le texte des cellules fusionnées et on lui applique le format "bold_right" défini plus haut:
        if col in fusion_cols:

            texte = zones_fusion[
                df.columns[col]
            ][0]

            worksheet.write(
                0,
                col,
                texte,
                formats["bold_right"]
            )

        # On parcourt les lignes de données et on leur applique le format "data_right" défini plus haut:
        for row in range(2, len(df) + 2):

            worksheet.write(
                row,
                col,
                df.iloc[row - 2, col],
                formats["data_right"]
            )

    # ------------------------
    # Mise en forme générale
    # ------------------------

    # On ajuste la largeur des colonnes à leur contenu:
    autofit_columns(
        df,
        worksheet,
        rotated_columns,
        wrapped_columns
    )

    # On formate la ligne TOTAL si celle-ci est présente:
    # On cherche une ligne TOTAL dans la première colonne

    premiere_colonne = df.columns[0]

    if (
        "TOTAL"
        in
        df[premiere_colonne].astype(str).values
    ):

        idx_total = df.index[
            df[premiere_colonne].astype(str)
            == "TOTAL"
        ][0]

        # Ligne Excel :
        # 0 = cellules fusionnées
        # 1 = en-têtes
        # 2 = première ligne de données

        ligne_total_excel = idx_total + 2

        for col_num, valeur in enumerate(
            df.iloc[idx_total]
        ):

            format_total = workbook.add_format({
                "bold": True,
                "top": 2
            })

            format_total_right = workbook.add_format({
                "bold": True,
                "top": 2,
                "right": 2
            })

            if col_num in cols_with_border:

                fmt = format_total_right

            else:

                fmt = format_total

            worksheet.write(
                ligne_total_excel,
                col_num,
                valeur,
                fmt
            )

    # On gèle les deux premières lignes dans Excel:
    worksheet.freeze_panes(2, 0)


    # On ajoute des filtres sur les en-têtes des colonnes dans Excel:
    worksheet.autofilter(
        1,
        0,
        len(df) + 1,
        len(df.columns) - 1
    )


##############################################################
# 4. Formatage excel pour exportation des feuilles "origine"
##############################################################
'''
On définit la fonction export_origine() qui exporte les dataframes "d'origine", soit les feuilles contenant toutes les données d'origine (champs 993, MMS IDs, et contenu séparé par colonne de chaque champ 993, i.e. bibliothèque, date, traitement, etc.) qui ne contiennent pas d'erreurs dans un fichier Excel formaté. La première feuille contient absolument tout, les deux autres contiennent les champs 993 en lien avec le catalogage, la troisième ceux en lien avec l'indexation. Le fichier est enregistré dans le dossier EXPORT_FOLDER, défini dans config.py. Si le dossier EXPORT_FOLDER n'existe pas, il est créé automatiquement, et il se trouve dans le dossier parent à celui contenant ce programme Python.
'''
def export_origine(writer, df, nom_feuille):

    # On transforme la dataframe en feuille Excel:
    df.to_excel(writer, sheet_name=nom_feuille, index=False)

    # On récupère le fichier et la feuille Excel pour pouvoir les formater:
    workbook = writer.book
    worksheet = writer.sheets[nom_feuille]

    # On récupère les formats créés précédemment:
    formats = create_formats(workbook)

    # On récupère l'indice de la colonne "MMS Id":
    col_mms = df.columns.get_loc("MMS Id")

    # On parcourt toutes les lignes de la feuille et on convertit chacune d'elles en texte (strings) afin d'éviter qu'Excel ne transforme les MMS IDs avec la notation scientifique:
    for row in range(1, len(df) + 1):
        worksheet.write(
            row,
            col_mms,
            str(df.iloc[row - 1, col_mms]),
            formats["text"]
        )

    # On parcout les en-têtes des colonnes et on leur applique le format "bold":
    for col_index, col_name in enumerate(df.columns):
        worksheet.write(
            0,
            col_index,
            col_name,
            formats["bold"]
        )



    # On ajuste la largeur des colonnes à leur contenu:
    autofit_columns(df, worksheet, [], [])

    # On gèle la première ligne dans Excel:
    worksheet.freeze_panes(1, 0)

    # On ajoute des filtres sur les en-têtes des colonnes dans Excel:
    worksheet.autofilter(
        0,
        0,
        len(df),
        len(df.columns) - 1
    )


########################################
# 5. Export Excel des champs corrects
########################################
'''
On définit la fonction export_corrects() qui exporte la liste de tous les champs 993 qui ne contiennent pas d'erreurs dans un fichier Excel formaté. Le fichier est enregistré dans le dossier EXPORT_FOLDER, défini dans config.py. Si le dossier EXPORT_FOLDER n'existe pas, il est créé automatiquement, et il se trouve dans le dossier parent à celui contenant ce programme Python.
'''
# La fonction prend comme arguments les différentes dataframes qui doivent être transformées en feuilles Excel:
def export_corrects(
        df_C_par_personne,
        df_C_par_biblio,
        df_I_par_personne,
        df_I_par_biblio,
        df_CENT,
        df_origine,
        df_cat,
        df_ind):

    # Définition du dossier courant:
    current_dir = os.path.dirname(__file__)

    # Définition du dossier parent au dossier courant:
    parent_dir = os.path.dirname(current_dir)

    # Définition du dossier dans lequel le fichier contenant les erreurs sera exporté:
    export_dir = os.path.join(
        parent_dir,
        EXPORT_FOLDER
    )

    # On crée le dossier export_dir si celui-ci n'existe pas encore:
    os.makedirs(export_dir, exist_ok=True)

    # On ajoute la date après le nom choisi dans config.py pour le nom du fichier contenant les champs 993 corrects (CORRECT_FILENAME):
    fname = CORRECT_FILENAME.format(
        date=f"{dt.datetime.today():%Y%m%d}"
    )

    # On crée le chemin d'exportation en joignant le chemin du dossier (export_dir) et le nom du fichier choisi (fname):
    path = os.path.join(export_dir, fname)

    # On définit nos différentes feuilles Excel et la façon dont elles doivent être formatées:
    feuilles = [
        {
            "df": df_C_par_personne,
            "nom": "C-nbr par personne",
            "zones": {
                "A": ("Mode d'acquisition", 5),
                "C": ("Type de traitement", 3),
                "01": ("Nombre CAT par mois", 5)
            },
            "rotated": [
                "Catalogage",
                "Notices sans inventaire",
                "Recatalogage",
                "Traitement pour une autre biblio",
                "Rien à préciser"
            ],
            "wrapped": ["Nombre total catalogage"],
            "borders": [
                "Nombre total catalogage",
                "Recatalogage",
                ("A", "avant"),
                ("C", "avant"),
                ("01", "avant")
            ]
        },

        {
            "df": df_C_par_biblio,
            "nom": "C-nbr par biblio",
            "zones": {
                "A": ("Mode d'acquisition", 5),
                "C": ("Type de traitement", 3),
                "01": ("Nombre CAT par mois", 5)
            },
            "rotated": [
                "Catalogage",
                "Notices sans inventaire",
                "Recatalogage",
                "Traitement pour une autre biblio",
                "Rien à préciser"
            ],
            "wrapped": ["Nombre total catalogage"],
            "borders": [
                "Nombre total catalogage",
                "Recatalogage",
                ("A", "avant"),
                ("C", "avant"),
                ("01", "avant")
            ]
        },

        {
            "df": df_I_par_personne,
            "nom": "I-nbr par personne",
            "zones": {
                "I": ("Type de traitement", 2),
                "01": ("Nombre indexation par mois", 6)
            },
            "rotated": [
                "Traitement pour une autre biblio",
                "Rien à préciser"
            ],
            "wrapped": ["Nombre total indexation"],
            "borders": [
                "Nombre total indexation",
                "Rien à préciser",
                ("01", "avant")
            ]
        },

        {
            "df": df_I_par_biblio,
            "nom": "I-nbr par biblio",
            "zones": {
                "I": ("Mode d'indexation", 2),
                "01": ("Nombre indexation par mois", 6)
            },
            "rotated": [
                "Traitement pour une autre biblio",
                "Rien à préciser"
            ],
            "wrapped": ["Nombre total indexation"],
            "borders": [
                "Nombre total indexation",
                "Rien à préciser",
                ("01", "avant")
            ]
        },

        {
            "df": df_CENT,
            "nom": "CENT",
            "zones": {
                "02": ("Nombres par mois", 6)
            },
            "rotated": [],
            "wrapped": ["Nombre total de livres traités"],
            "borders": [
                "Nombre total de livres traités",
                ("02", "avant")
            ]
        }
    ]

    # Enregistrement du fichier Excel proprement dit, avec le formatage requis:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:

        for feuille in feuilles:

            df_export = feuille["df"].copy()

            if len(df_export) > 1:

                ligne_total = {}

                premiere_colonne = df_export.columns[0]

                for col in df_export.columns:

                    if col == premiere_colonne:

                        ligne_total[col] = "TOTAL"

                    else:

                        ligne_total[col] = (
                            df_export[col].sum()
                        )

                df_export = pd.concat(
                    [
                        df_export,
                        pd.DataFrame([ligne_total])
                    ],
                    ignore_index=True
                )

            formater_feuille_statistique(
                writer,
                df_export,
                feuille["nom"],
                feuille["zones"],
                feuille["rotated"],
                feuille["wrapped"],
                feuille["borders"]
            )

        export_origine(writer, df_origine, "Origine")
        export_origine(writer, df_ind, "Origine IND")
        export_origine(writer, df_cat, "Origine CAT")



###############################
# 6. Export Excel des erreurs
###############################
'''
On définit la fonction export_errors() qui exporte la liste de toutes les erreurs recensées dans un fichier Excel formaté. Le fichier est enregistré dans le dossier EXPORT_FOLDER, défini dans config.py. Si le dossier EXPORT_FOLDER n'existe pas, il est créé automatiquement, et il se trouve dans le dossier parent à celui contenant ce programme Python.
'''
def export_errors(df):

    # Définition du dossier courant:
    current_dir = os.path.dirname(__file__)

    # Définition du dossier parent au dossier courant:
    parent_dir = os.path.dirname(current_dir)

    # Définition du dossier dans lequel le fichier contenant les erreurs sera exporté:
    export_dir = os.path.join(
        parent_dir,
        EXPORT_FOLDER
    )

    # On crée le dossier export_dir si celui-ci n'existe pas encore:
    os.makedirs(export_dir, exist_ok=True)

    # On ajoute la date après le nom choisi dans config.py pour le nom du fichier contenant les erreurs détectées (ERROR_FILENAME):
    fname = ERROR_FILENAME.format(
        date=f"{dt.datetime.today():%Y%m%d}"
    )

    # On crée le chemin d'exportation en joignant le chemin du dossier (export_dir) et le nom du fichier choisi (fname):
    path = os.path.join(export_dir, fname)

    # Enregistrement du fichier Excel proprement dit, avec le formatage requis:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # La feuille Excel est nommée "Erreurs"
        df.to_excel(writer, sheet_name="Erreurs")
        ws = writer.sheets["Erreurs"]

        # La ligne d'en-tête sera en gras:
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # La ligne d'en-tête contiendra des filtres qui permettront d'isoler les erreurs qui nous intéressent, si besoin est:
        ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=ws.max_column).coordinate}"

        # Ajustement de la largeur des colonnes à leur contenu:
        for col in ws.columns:
            max_len = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2


##################################################
# Nouvelle fonction pour les fichiers par personne
##################################################

def export_par_personne(person_files):

    current_dir = os.path.dirname(__file__)
    parent_dir = os.path.dirname(current_dir)

    export_dir = os.path.join(
        parent_dir,
        EXPORT_FOLDER,
        "Par_personne"
    )

    os.makedirs(export_dir, exist_ok=True)

    for personne, data in person_files.items():

        nom_personne = (
            str(personne)
            .replace("/", "_")
            .replace("\\", "_")
        )

        fname = (
            f"{nom_personne}_"
            f"{dt.datetime.today():%Y%m%d}.xlsx"
        )

        path = os.path.join(
            export_dir,
            fname
        )

        with pd.ExcelWriter(
            path,
            engine="xlsxwriter"
        ) as writer:

            # ==================================
            # FEUILLE DETAIL
            # ==================================

            export_origine(
                writer,
                data["detail"],
                "Détail"
            )

            # ==================================
            # FEUILLE C-nbr par biblio
            # ==================================

            if data["cat_biblio"] is not None:

                df_cat_biblio = data["cat_biblio"].copy()

                if len(df_cat_biblio) > 1:

                    ligne_total = {}

                    for col in df_cat_biblio.columns:

                        if col == "Bibliothèque":

                            ligne_total[col] = "TOTAL"

                        else:

                            ligne_total[col] = (
                                df_cat_biblio[col].sum()
                            )

                    df_cat_biblio = pd.concat(
                        [
                            df_cat_biblio,
                            pd.DataFrame([ligne_total])
                        ],
                        ignore_index=True
                    )

                formater_feuille_statistique(
                    writer,
                    df_cat_biblio,
                    "Catalogage",
                    {
                        "A": (
                            "Mode d'acquisition",
                            5
                        ),
                        "C": (
                            "Type de traitement",
                            3
                        ),
                        "01": (
                            "Nombre CAT par mois",
                            5
                        )
                    },
                    [
                        "Catalogage",
                        "Notices sans inventaire",
                        "Recatalogage",
                        "Traitement pour une autre biblio",
                        "Rien à préciser"
                    ],
                    [
                        "Nombre total catalogage"
                    ],
                    [
                        "Nombre total catalogage",
                        "Recatalogage",
                        ("A", "avant"),
                        ("C", "avant"),
                        ("01", "avant")
                    ]
                )

            # ==================================
            # FEUILLE I-nbr par biblio
            # ==================================

            if data["ind_biblio"] is not None:

                df_ind_biblio = data["ind_biblio"].copy()

                if len(df_ind_biblio) > 1:

                    ligne_total = {}

                    for col in df_ind_biblio.columns:

                        if col == "Bibliothèque":

                            ligne_total[col] = "TOTAL"

                        else:

                            ligne_total[col] = (
                                df_ind_biblio[col].sum()
                            )

                    df_ind_biblio = pd.concat(
                        [
                            df_ind_biblio,
                            pd.DataFrame([ligne_total])
                        ],
                        ignore_index=True
                    )

                formater_feuille_statistique(
                    writer,
                    df_ind_biblio,
                    "Indexation",
                    {
                        "I": (
                            "Mode d'indexation",
                            2
                        ),
                        "01": (
                            "Nombre indexation par mois",
                            6
                        )
                    },
                    [
                        "Traitement pour une autre biblio",
                        "Rien à préciser"
                    ],
                    [
                        "Nombre total indexation"
                    ],
                    [
                        "Nombre total indexation",
                        "Rien à préciser",
                        ("01", "avant")
                    ]
                )

            # ==================================
            # FEUILLE DOMAINES
            # ==================================

            if (
                data["domaines"] is not None
                and not data["domaines"].empty
            ):

                df_dom = data["domaines"].copy()

                df_dom.to_excel(
                    writer,
                    sheet_name="Domaines",
                    index=False
                )

                ws = writer.sheets["Domaines"]

                # En-têtes en gras

                formats = create_formats(writer.book)

                for col_num, col_name in enumerate(data["domaines"].columns):
                    ws.write(
                        0,
                        col_num,
                        col_name,
                        formats["bold"]
                    )

                # --------------------------------
                # Ligne TOTAL
                # --------------------------------

                ligne_total = len(df_dom) + 2

                ws.write(ligne_total, 0, "TOTAL")

                for col in range(1, len(df_dom.columns)):

                    total = df_dom.iloc[:, col].sum()

                    bold = writer.book.add_format({
                        "bold": True
                    })

                    ws.write(
                        ligne_total,
                        col,
                        total,
                        bold
                    )

                # --------------------------------
                # Bloc de texte explicatif
                # --------------------------------

                note_format = writer.book.add_format({
                    "text_wrap": True,
                    "bg_color": "#F2F2F2",
                    "border": 1
                })


                # Positionnement de la remarque une colonne après la dernière colonne du tableau:
                col_depart = len(df_dom.columns) + 1

                ws.merge_range(
                    1,              # ligne
                    col_depart,     # Colonne où la remarque apparaîtra
                    7,              # ligne 7
                    col_depart+4,   # Colonne où la remarque se termine
                    (
                        "Remarque :\n"
                        "Une même notice peut être "
                        "associée à plusieurs domaines. "
                        "Les totaux affichés ici ne correspondent "
                        "donc pas au nombre réel de notices traitées.\n"
                        "Pour le nombre total de notices traitées, "
                        "merci de vous réferrer aux feuilles "
                        "'Catalogage' ou 'Indexation'."
                    ),
                    note_format
                )

                autofit_columns(
                    df_dom,
                    ws,
                    [],
                    []
                )

                ws.freeze_panes(1, 0)

                ws.autofilter(
                    0,
                    0,
                    len(df_dom),
                    len(df_dom.columns) - 1
                )


########################################################
# Nouvelle fonction pour obtenir un fichier par biblio
########################################################

def export_par_bibliotheque(bibli_files):

    current_dir = os.path.dirname(__file__)
    parent_dir = os.path.dirname(current_dir)

    export_dir = os.path.join(
        parent_dir,
        EXPORT_FOLDER,
        "Par_bibliotheque"
    )

    os.makedirs(export_dir, exist_ok=True)

    for biblio, data in bibli_files.items():

        nom_biblio = (
            str(biblio)
            .replace("/", "_")
            .replace("\\", "_")
        )

        fname = (
            f"{nom_biblio}_"
            f"{dt.datetime.today():%Y%m%d}.xlsx"
        )

        path = os.path.join(
            export_dir,
            fname
        )

        with pd.ExcelWriter(
            path,
            engine="xlsxwriter"
        ) as writer:

            # ==================================
            # FEUILLE DETAIL
            # ==================================

            export_origine(
                writer,
                data["detail"],
                "Détail"
            )

            # ==================================
            # FEUILLE CATALOGAGE
            # ==================================

            if data["cat"] is not None:

                df_cat_export = data["cat"].copy()

                # Ajout d'une ligne TOTAL uniquement
                # s'il y a plusieurs personnes

                if len(df_cat_export) > 1:

                    ligne_total = {}

                    for col in df_cat_export.columns:

                        if col == "Personne":

                            ligne_total[col] = "TOTAL"

                        else:

                            ligne_total[col] = (
                                df_cat_export[col].sum()
                            )

                    df_cat_export = pd.concat(
                        [
                            df_cat_export,
                            pd.DataFrame([ligne_total])
                        ],
                        ignore_index=True
                    )

                formater_feuille_statistique(
                    writer,
                    df_cat_export,
                    "C-nbr par personne",
                    {
                        "A": (
                            "Mode d'acquisition",
                            5
                        ),
                        "C": (
                            "Type de traitement",
                            3
                        ),
                        "01": (
                            "Nombre CAT par mois",
                            5
                        )
                    },
                    [
                        "Catalogage",
                        "Notices sans inventaire",
                        "Recatalogage",
                        "Traitement pour une autre biblio",
                        "Rien à préciser"
                    ],
                    [
                        "Nombre total catalogage"
                    ],
                    [
                        "Nombre total catalogage",
                        "Recatalogage",
                        ("A", "avant"),
                        ("C", "avant"),
                        ("01", "avant")
                    ]
                )

                # Ajout d'une remarque expliquant que toutes les personnes présentes dans le fichier de la bibliothèque ne font pas forcément partie du personnel de la bibliothèque:
                ws = writer.sheets["C-nbr par personne"]

                note_format = writer.book.add_format({
                    "text_wrap": True,
                    "bg_color": "#F2F2F2",
                    "border": 1
                })

                # Le bloc commence 1 colonne après la dernière colonne du tableau
                col_depart = len(df_cat_export.columns) + 1

                ws.merge_range(
                    2,
                    col_depart,
                    5,
                    col_depart + 7, # largeur du bloc de texte
                    (
                        "Remarque :\n"
                        "Les personnes figurant dans cette feuille "
                        "ne font pas nécessairement partie du personnel "
                        "de la bibliothèque. Toute personne ayant effectué "
                        "du catalogage pour cette bibliothèque est prise en compte."
                    ),
                    note_format
                )



            # ==================================
            # FEUILLE C-nbr par biblio
            # ==================================

            if (
                biblio == "COLLUNI"
                and
                data["cat_biblio"] is not None
            ):

                df_cat_biblio = (
                    data["cat_biblio"]
                    .copy()
                )

                if len(df_cat_biblio) > 1:

                    ligne_total = {}

                    for col in df_cat_biblio.columns:

                        if col == "Bibliothèques_COLLUNI":

                            ligne_total[col] = "TOTAL"

                        else:

                            ligne_total[col] = (
                                df_cat_biblio[col]
                                .sum()
                            )

                    df_cat_biblio = pd.concat(
                        [
                            df_cat_biblio,
                            pd.DataFrame([ligne_total])
                        ],
                        ignore_index=True
                    )

                formater_feuille_statistique(
                    writer,
                    df_cat_biblio,
                    "C-nbr par biblio",
                    {
                        "A": (
                            "Mode d'acquisition",
                            5
                        ),
                        "C": (
                            "Type de traitement",
                            3
                        ),
                        "01": (
                            "Nombre CAT par mois",
                            5
                        )
                    },
                    [
                        "Catalogage",
                        "Notices sans inventaire",
                        "Recatalogage",
                        "Traitement pour une autre biblio",
                        "Rien à préciser"
                    ],
                    [
                        "Nombre total catalogage"
                    ],
                    [
                        "Nombre total catalogage",
                        "Recatalogage",
                        ("A", "avant"),
                        ("C", "avant"),
                        ("01", "avant")
                    ]
                )


            # ==================================
            # FEUILLE INDEXATION
            # ==================================

            if data["ind"] is not None:

                df_ind_export = data["ind"].copy()

                # Ajout d'une ligne TOTAL uniquement
                # s'il y a plusieurs personnes

                if len(df_ind_export) > 1:

                    ligne_total = {}

                    for col in df_ind_export.columns:

                        if col == "Personne":

                            ligne_total[col] = "TOTAL"

                        else:

                            ligne_total[col] = (
                                df_ind_export[col].sum()
                            )


                    df_ind_export = pd.concat(
                        [
                            df_ind_export,
                            pd.DataFrame([ligne_total])
                        ],
                        ignore_index=True
                    )


                formater_feuille_statistique(
                    writer,
                    df_ind_export,
                    "I-nbr par personne",
                    {
                        "I": (
                            "Mode d'indexation",
                            2
                        ),
                        "01": (
                            "Nombre indexation par mois",
                            6
                        )
                    },
                    [
                        "Traitement pour une autre biblio",
                        "Rien à préciser"
                    ],
                    [
                        "Nombre total indexation"
                    ],
                    [
                        "Nombre total indexation",
                        "Rien à préciser",
                        ("01", "avant")
                    ]
                )

                # Ajout d'une remarque expliquant que toutes les personnes présentes dans le fichier de la bibliothèque ne font pas forcément partie du personnel de la bibliothèque:
                ws = writer.sheets["I-nbr par personne"]

                note_format = writer.book.add_format({
                    "text_wrap": True,
                    "bg_color": "#F2F2F2",
                    "border": 1
                })

                col_depart = len(df_ind_export.columns) + 1

                ws.merge_range(
                    2,
                    col_depart,
                    5,
                    col_depart + 7, # largeur du bloc de texte
                    (
                        "Remarque :\n"
                        "Les personnes figurant dans cette feuille "
                        "ne font pas nécessairement partie du personnel "
                        "de la bibliothèque. Toute personne ayant effectué "
                        "de l'indexation pour cette bibliothèque est prise en compte."
                    ),
                    note_format
                )


            # ==================================
            # FEUILLE I-nbr par biblio
            # ==================================

            if (
                biblio == "COLLUNI"
                and
                data["ind_biblio"] is not None
            ):

                df_ind_biblio = (
                    data["ind_biblio"]
                    .copy()
                )

                if len(df_ind_biblio) > 1:

                    ligne_total = {}

                    for col in df_ind_biblio.columns:

                        if col == "Bibliothèques_COLLUNI":

                            ligne_total[col] = "TOTAL"

                        else:

                            ligne_total[col] = (
                                df_ind_biblio[col]
                                .sum()
                            )

                    df_ind_biblio = pd.concat(
                        [
                            df_ind_biblio,
                            pd.DataFrame([ligne_total])
                        ],
                        ignore_index=True
                    )

                formater_feuille_statistique(
                    writer,
                    df_ind_biblio,
                    "I-nbr par biblio",
                    {
                        "I": (
                            "Mode d'indexation",
                            2
                        ),
                        "01": (
                            "Nombre indexation par mois",
                            6
                        )
                    },
                    [
                        "Traitement pour une autre biblio",
                        "Rien à préciser"
                    ],
                    [
                        "Nombre total indexation"
                    ],
                    [
                        "Nombre total indexation",
                        "Rien à préciser",
                        ("01", "avant")
                    ]
                )

            # ==================================
            # FEUILLE DOMAINES
            # ==================================

            if (
                data["domaines"] is not None
                and
                not data["domaines"].empty
            ):

                df_dom = data["domaines"].copy()

                df_dom.to_excel(
                    writer,
                    sheet_name="Domaines",
                    index=False
                )

                ws = writer.sheets["Domaines"]

                # -----------------------------
                # En-têtes en gras
                # -----------------------------

                formats = create_formats(
                    writer.book
                )

                for col_num, col_name in enumerate(
                        df_dom.columns):

                    ws.write(
                        0,
                        col_num,
                        col_name,
                        formats["bold"]
                    )

                # -----------------------------
                # Ligne TOTAL
                # -----------------------------

                ligne_total = len(df_dom) + 2

                bold = writer.book.add_format({
                    "bold": True
                })

                ws.write(
                    ligne_total,
                    0,
                    "TOTAL",
                    bold
                )

                for col in range(
                        1,
                        len(df_dom.columns)
                ):

                    total = (
                        df_dom
                        .iloc[:, col]
                        .sum()
                    )

                    ws.write(
                        ligne_total,
                        col,
                        total,
                        bold
                    )

                # --------------------------------
                # Bloc de texte explicatif
                # --------------------------------

                note_format = writer.book.add_format({
                    "text_wrap": True,
                    "bg_color": "#F2F2F2",
                    "border": 1
                })

                # Positionnement du texte une colonne après la dernière colonne du tableau contenant qqch:
                col_depart = len(df_dom.columns) + 1

                ws.merge_range(
                    1,
                    col_depart,
                    7,
                    col_depart + 4,
                    (
                        "Remarque :\n"
                        "Une même notice peut être "
                        "associée à plusieurs domaines. "
                        "Les totaux affichés ici ne correspondent "
                        "donc pas au nombre réel de notices traitées.\n"
                        "Pour le nombre total de notices traitées, "
                        "merci de vous référer aux feuilles "
                        "'Catalogage' ou 'Indexation'."
                    ),
                    note_format
                )

                autofit_columns(
                    df_dom,
                    ws,
                    [],
                    []
                )

                ws.freeze_panes(1, 0)

                ws.autofilter(
                    0,
                    0,
                    len(df_dom),
                    len(df_dom.columns) - 1
                )



