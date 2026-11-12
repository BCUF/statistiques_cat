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

from excel_helper import ExcelHelper

class ExcelExporter:

    def __init__(self, export_folder=EXPORT_FOLDER):
        self.export_dir= os.path.abspath(export_folder)

    ########################################
    # 5. Export Excel des champs corrects
    ########################################
    '''
    On définit la fonction export_corrects() qui exporte la liste de tous les champs 993 qui ne contiennent pas d'erreurs dans un fichier Excel formaté. Le fichier est enregistré dans le dossier EXPORT_FOLDER, défini dans config.py. Si le dossier EXPORT_FOLDER n'existe pas, il est créé automatiquement, et il se trouve dans le dossier parent à celui contenant ce programme Python.
    '''
    # La fonction prend comme arguments les différentes dataframes qui doivent être transformées en feuilles Excel:
    def export_corrects(
            self,
            df_C_par_personne,
            df_C_par_biblio,
            df_I_par_personne,
            df_I_par_biblio,
            df_CENT,
            df_origine,
            df_cat,
            df_ind):

        # On crée le dossier export_dir si celui-ci n'existe pas encore:
        os.makedirs(self.export_dir, exist_ok=True)

        # On ajoute la date après le nom choisi dans config.py pour le nom du fichier contenant les champs 993 corrects (CORRECT_FILENAME):
        fname = CORRECT_FILENAME.format(
            date=f"{dt.datetime.today():%Y%m%d}"
        )

        # On crée le chemin d'exportation en joignant le chemin du dossier (self.export_dir) et le nom du fichier choisi (fname):
        path = os.path.join(self.export_dir, fname)

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

                ExcelHelper.formater_feuille_statistique(
                    writer,
                    df_export,
                    feuille["nom"],
                    feuille["zones"],
                    feuille["rotated"],
                    feuille["wrapped"],
                    feuille["borders"]
                )

            ExcelHelper.export_origine(writer, df_origine, "Origine")
            ExcelHelper.export_origine(writer, df_cat, "Origine CAT")
            ExcelHelper.export_origine(writer, df_ind, "Origine IND")




    ###############################
    # 6. Export Excel des erreurs
    ###############################
    '''
    On définit la fonction export_errors() qui exporte la liste de toutes les erreurs recensées dans un fichier Excel formaté. Le fichier est enregistré dans le dossier EXPORT_FOLDER, défini dans config.py. Si le dossier EXPORT_FOLDER n'existe pas, il est créé automatiquement, et il se trouve dans le dossier parent à celui contenant ce programme Python.
    '''
    def export_errors(self, df):

        # On crée le dossier self.export_dir si celui-ci n'existe pas encore:
        os.makedirs(self.export_dir, exist_ok=True)

        # On ajoute la date après le nom choisi dans config.py pour le nom du fichier contenant les erreurs détectées (ERROR_FILENAME):
        fname = ERROR_FILENAME.format(
            date=f"{dt.datetime.today():%Y%m%d}"
        )

        # On crée le chemin d'exportation en joignant le chemin du dossier (self.export_dir) et le nom du fichier choisi (fname):
        path = os.path.join(self.export_dir, fname)

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

    def export_par_personne(self, person_files):


        par_personne_export_dir = os.path.join(
            self.export_dir,
            "Par_personne"
        )

        os.makedirs(par_personne_export_dir, exist_ok=True)

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
                par_personne_export_dir,
                fname
            )

            with pd.ExcelWriter(
                path,
                engine="xlsxwriter"
            ) as writer:

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

                    ExcelHelper.formater_feuille_statistique(
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

                    ExcelHelper.formater_feuille_statistique(
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

                    formats = ExcelHelper.create_formats(writer.book)

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

                    ExcelHelper.autofit_columns(
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

                # ==================================
                # FEUILLE DETAIL
                # ==================================

                ExcelHelper.export_origine(
                    writer,
                    data["detail"],
                    "Détail"
                )


    ########################################################
    # Nouvelle fonction pour obtenir un fichier par biblio
    ########################################################

    def export_par_bibliotheque(self, bibli_files):

        os.makedirs(self.export_dir, exist_ok=True)

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
                self.export_dir,
                fname
            )

            with pd.ExcelWriter(
                path,
                engine="xlsxwriter"
            ) as writer:


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

                    ExcelHelper.formater_feuille_statistique(
                        writer,
                        df_cat_export,
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
                        ],
                        titre="Catalogage par personne",
                        remarque=(
                            "Remarque :\n"
                            "Les personnes figurant dans cette feuille "
                            "ne font pas nécessairement partie du personnel "
                            "de la bibliothèque. Toute personne ayant effectué "
                            "du catalogage pour cette bibliothèque est prise en compte."
                        )
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

                    ExcelHelper.formater_feuille_statistique(
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


                    ExcelHelper.formater_feuille_statistique(
                        writer,
                        df_ind_export,
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
                        ],
                            titre="Indexation par personne",
                            remarque=(
                                "Remarque :\n"
                                "Les personnes figurant dans cette feuille "
                                "ne font pas nécessairement partie du personnel "
                                "de la bibliothèque. Toute personne ayant effectué "
                                "de l'indexation pour cette bibliothèque est prise en compte."
                            )
                        
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

                    ExcelHelper.formater_feuille_statistique(
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

                    formats = ExcelHelper.create_formats(
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

                    ExcelHelper.autofit_columns(
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

                # ==================================
                # FEUILLE DETAIL
                # ==================================

                ExcelHelper.export_origine(
                    writer,
                    data["detail"],
                    "Détail"
                )
