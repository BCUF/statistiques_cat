# Ce fichier contient les paramètres modifiables du programme. On les réunit dans un fichier séparé afin de se faciliter le travail à l'avenir, lorsqu'il faudra par exemple lancer le code en 2027 (il suffira alors de changer la valeur de la variable YEAR dans ce fichier), ou si l'on souhaite changer les noms de fichiers utilisés (INPUT_FILENAME) et exportés (ERROR_FILENAME et CORRECT_FILENAME).

# Nom du dossier dans lequel se trouve le fichier exporté depuis Alma:
INPUT_FOLDER = "Données"

# Nom du dossier dans lequel les fichiers Excel seront exportés. Si celui-ci n'existe pas, il sera créé par notre programme:
EXPORT_FOLDER = "Exports"

# Nom du fichier natif, exporté depuis Alma, qui contient les colonnes "MMS Id" et "993 - Local Param 14" (i.e. les champs 993):
INPUT_FILENAME = "champs_990_993_all_{date}.xlsx"

# Nom du fichier qui contiendra les erreurs repérées:
ERROR_FILENAME = "erreurs_{date}.xlsx"

# Nom du fichier qui contiendra les informations relatives aux champs corrects, qui ne contiennent pas d'erreurs:
CORRECT_FILENAME = "champs_corrects_separes_{date}.xlsx"

# Année en cours: à changer chaque année:
YEAR = "2026"

# Regex pour le mois et le jour:
MONTH_PATTERN = r"[01]\d"
DAY_PATTERN = r"[0-3]\d"

# Regex pour la date entière:
DATE_PATTERN = rf"{YEAR}{MONTH_PATTERN}{DAY_PATTERN}"

# Regex pour la date (qu'on utilise une fois dans transformations_df.py):
DATE_SPLIT_PATTERN = (
    rf'({YEAR})'
    rf'({MONTH_PATTERN})'
    rf'({DAY_PATTERN})'
)

# Regex pour le catalogage:
CATALOGAGE_PATTERN = (
    rf'\$\$a\s+[A-Z2_]{{5}}/'
    rf'{DATE_PATTERN}/'
    rf'[a-z]{{3}}/'
    rf'[CRN][A_][ASZDGELFOPK_][CRIV]'
)

# Regex pour l'indexation:
INDEXATION_PATTERN = (
    rf'\$\$a\s+[A-Z2_]{{5}}/'
    rf'{DATE_PATTERN}/'
    rf'[a-z]{{3}}/'
    rf'I[A_][IJN]_'
)

# Regex permettant de séparer les champs 993 dans leurs différentes composantes dans transformations_df.py:
CAT_IND_PATTERN = (
    rf'\$\$a\s+([A-Z2_]{{5}})'
    rf'/({DATE_PATTERN})'
    rf'/([a-z]{{3}})'
    rf'/([CRNI][A_][ASZDGELFIJNOPK_][CRIV_])'
)

# Regex pour les différents traitements et modes d'acquisition possibles (TRAITEMENT_PATTERN est utilisé dans transformations_df.py):
TRAITEMENT_1_CODES = "CRNI"
TRAITEMENT_2_CODES = "A_"
TRAITEMENT_3_CODES = "ASZDGELFIJNOPK_"
TRAITEMENT_4_CODES = "CRIV_"

TRAITEMENT_PATTERN = (
    rf'([{TRAITEMENT_1_CODES}])'
    rf'([{TRAITEMENT_2_CODES}])'
    rf'([{TRAITEMENT_3_CODES}])'
    rf'([{TRAITEMENT_4_CODES}])'
)

# Nombre de champs 993 maximal: on estime que ce nombre ne dépassera jamais 50 pour un seul MMS ID
MAX_993_FIELDS = 50

# Dictionnaire des domaines
CDD_MAPPING = {
    "001": "Friburgensia",
    "004": "Informatique",
    "0081": "CD Blues",
    "0082": "CD Deutschsprachige Liedermacher",
    "0083": "CD Chanson française",
    "0084": "CD Classique",
    "0085": "CD World",
    "0086": "CD Jazz",
    "0087": "CD Livres audio",
    "009": "DVD Films de fiction",
    "020": "Bibliothéconomie",
    "070": "DVD Films documentaires",
    "306": "Anthropologie, ethnologie, traditions",
    "930": "Antiquité, archéologie",
    "720": "Architecture",
    "791": "Arts du spectacle, cinéma",
    "700": "Arts (généralités), peinture, sculpture",
    "570": "Biologie",
    "540": "Chimie",
    "302": "Communication, médias",
    "641": "Cuisine, alimentation",
    "340": "Droit",
    "577": "Ecologie, environnement",
    "330": "Economie",
    "370": "Education, pédagogie",
    "576": "Génétique, évolution",
    "910": "Géographie, voyages",
    "658": "Gestion, management",
    "900": "Histoire",
    "790": "Loisirs",
    "510": "Mathématiques",
    "610": "Médecine, santé",
    "780": "Musique",
    "100": "Philosophie",
    "770": "Photographie",
    "530": "Physique",
    "320": "Politique",
    "150": "Psychologie",
    "600": "Sciences appliquées, technologie",
    "550": "Sciences de la Terre",
    "500": "Sciences naturelles",
    "301": "Sociologie",
    "796": "Sport",
    "200": "Théologie, religions",
    "418": "Didactique des langues",
    "741": "Genre: bandes dessinées",
    "920": "Genre: biographies",
    "808": "Genre: romans",
    "820": "Langues et littératures anglaises et américaines",
    "890": "Langues et littératures diverses",
    "860": "Langues et littératures espagnoles et portugaises",
    "840": "Langues et littératures françaises",
    "830": "Langues et littératures germaniques",
    "880": "Langues et littératures grecques",
    "850": "Langues et littératures italiennes",
    "870": "Langues et littératures latines",
    "859": "Langues et littératures rhéto-romanes",
    "891": "Langues et littératures slaves",
    "410": "Linguistique générale",
    "800": "Littérature générale et comparée",
    "899": "Littérature suisse"
}

# Personnel de chaque bibliothèque: créé initialement pour trier les personnes pour chaque bibliothèque, mais trop difficile à maintenir sur le long terme. A la place, une remarque indiquera dans chaque fichier que toutes les personnes présentes dans la liste ne font pas nécessairement partie du personnel de la bibliothèque, mais qu'elles ont apporté une aide ponctuelle à la bibliothèque concernée.
BIBLIO_PERSONNES = {
    "AEF": [
        "fcl",
        "nsc",
        "rya"
    ],

    "AVF": [
        "fcl"
    ],

    "BEF": [
        "ell"
    ],

    "BFD": [
        "cho",
        "clg",
        "dme",
        "mco",
        "mib"
    ],

    "BHAP": [
        "clc",
        "dch",
        "lem",
        "mdo",
        "mne",
        "msc"
    ],

    "BHT": [
        "ddu",
        "fsp",
        "lem",
        "mdo",
        "mjo",
        "nci",
        "pvo",
        "vba"
    ],

    "BLE": [
        "asa",
        "mja",
        "mso"
    ],

    "BLL": [
        "acb",
        "cms",
        "mso",
        "mtl"
    ],

    "BP2": [
        "gmi",
        "imc",
        "jde",
        "jle",
        "mdo",
        "msc",
        "mtl",
        "pbr",
        "pvo"
    ],

    "CENT": [
        "aal",
        "aas",
        "acb",
        "asa",
        "car",
        "cat",
        "cbi",
        "clc",
        "cms",
        "cri",
        "dch",
        "ddu",
        "dgu",
        "eob",
        "ero",
        "fcl",
        "fcr",
        "fsp",
        "hde",
        "jga",
        "jgr",
        "lbu",
        "lco",
        "lem",
        "mdo",
        "mja",
        "mjo",
        "mne",
        "mre",
        "msc",
        "mso",
        "mtl",
        "pbr",
        "pvo",
        "ser",
        "sle",
        "slo",
        "vba",
        "zpe"
    ],

    "COF": [
        "cat"
    ],

    "DMDP": [
        "aoe",
        "bbh",
        "mba"
    ],

    "DOKPE": [
        "sbr"
    ],

    "EOC": [
        "aas",
        "lem"
    ],

    "IFF": [
        "cho",
        "dme",
        "mco"
    ],

    "IIEDH": [
        "dme",
        "mib"
    ],

    "IPC": [
        "dbl",
        "ell"
    ],

    "MAH": [
        "fcl"
    ],

    "MAR": [
        "fcl"
    ],

    "MFDP": [
        "aca",
        "gle",
        "lcu",
        "sgo"
    ],

    "MUS": [
        "lem"
    ],

    "PSPE": [
        "ade",
        "cst",
        "dbl"
    ],

    "SAEF": [
        "fcl"
    ],

    "SBC": [
        "fcl"
    ],

    "SCANT": [
        "clc",
        "eob",
        "mib",
        "mne"
    ],

    "SDU": [
        "cho",
        "mbi"
    ],

    "STS": [
        "imc",
        "jde",
        "pbr"
    ],

    "TCF": [
        "api"
    ],

    "VCR": [
        "fcl"
    ]
}

#################################################
# Mode de génération des fichiers
#################################################

# Remplacer MODE_EXPORT = "menu" par MODE_EXPORT = "auto" si on ne souhaite finalement plus afficher le menu au lancement du programme
MODE_EXPORT = "menu"
