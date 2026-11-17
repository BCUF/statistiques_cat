"""Defaults and shared settings used by the program."""
import os
from typing import NamedTuple

# Nom du dossier dans lequel les fichiers Excel seront exportés. Si celui-ci n'existe pas, il sera créé par notre programme:
EXPORT_FOLDER = os.getenv("STAT_EXPORT_FOLDER", "../Exports")

# Nom du fichier qui contiendra les erreurs repérées:
ERROR_FILENAME = "erreurs_{date}.xlsx"

# Nom du fichier qui contiendra les informations relatives aux champs corrects, qui ne contiennent pas d'erreurs:
CORRECT_FILENAME = "champs_corrects_separes_{date}.xlsx"

# Année utilisée par défaut si aucune année n'est choisie dans le CLI ou le GUI:
YEAR = "2026"

# Regex pour le mois et le jour:
MONTH_PATTERN = r"[01]\d"
DAY_PATTERN = r"[0-3]\d"


class YearPatterns(NamedTuple):
    date_pattern: str
    date_split_pattern: str
    catalogage_pattern: str
    indexation_pattern: str
    cat_ind_pattern: str


def validate_year(year: str) -> str:
    if len(year) != 4 or not year.isascii() or not year.isdigit():
        raise ValueError("L'année doit être composée de quatre chiffres.")
    return year


def get_year_patterns(year: str) -> YearPatterns:
    year = validate_year(year)
    date_pattern = rf"{year}{MONTH_PATTERN}{DAY_PATTERN}"
    date_split_pattern = (
        rf'({year})'
        rf'({MONTH_PATTERN})'
        rf'({DAY_PATTERN})'
    )
    catalogage_pattern = (
        rf'\$\$a\s+[A-Z2_]{{5}}/'
        rf'{date_pattern}/'
        rf'[a-z]{{3}}/'
        rf'[CRN][A_][ASZDGELFOPK_][CRIV]'
    )
    indexation_pattern = (
        rf'\$\$a\s+[A-Z2_]{{5}}/'
        rf'{date_pattern}/'
        rf'[a-z]{{3}}/'
        rf'I[A_][IJN]_'
    )
    cat_ind_pattern = (
        rf'\$\$a\s+([A-Z2_]{{5}})'
        rf'/({date_pattern})'
        rf'/([a-z]{{3}})'
        rf'/([CRNI][A_][ASZDGELFIJNOPK_][CRIV_])'
    )
    return YearPatterns(
        date_pattern,
        date_split_pattern,
        catalogage_pattern,
        indexation_pattern,
        cat_ind_pattern,
    )


_DEFAULT_YEAR_PATTERNS = get_year_patterns(YEAR)
DATE_PATTERN = _DEFAULT_YEAR_PATTERNS.date_pattern
DATE_SPLIT_PATTERN = _DEFAULT_YEAR_PATTERNS.date_split_pattern
CATALOGAGE_PATTERN = _DEFAULT_YEAR_PATTERNS.catalogage_pattern
INDEXATION_PATTERN = _DEFAULT_YEAR_PATTERNS.indexation_pattern
CAT_IND_PATTERN = _DEFAULT_YEAR_PATTERNS.cat_ind_pattern

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

