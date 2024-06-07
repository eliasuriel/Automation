#Default values
DEFAULT_EXIT = 1
DEFAULT_CONTINUE = 0
DEFAULT_DEFAULT = 2
DEFAULT_YEAR = "2024"
DEFAULT_WK_HOURS = 8

PRIME_PATH = "/App/Prime05-abril-2024.xlsx"
OPX_PATH   = "/App/opx_05-abril-2024.xlsx"

#"C:\Users\VEU1GA\Documents\Visual_Studio\App\Prime05-abril-2024.xlsx"
DEFAULT_OUT_PATH   = "\App"

#Column headers
FIRST_COLUMN_NAME_PRIME = "Resource Name"

#Output File
OUTPUT_FILE_PREFIX = "Output-"
SHEET_2_NAME = "Associate List "
TABLE_START_COLUMN = 2
TABLE_START_ROW = 2
TABLE_STYLE = "TableStyleMedium9"
OUTPUT_HEADERS = [
    "General Project",
    "Month",
    "Component",
    "Associate Name",
    "OPX Project",
    "Prime Project",
    "OPX Hours",
    "Prime Hours",
    "Relation"
]


# Button names
choices = ["Continuar", "Salir"]
initial_choices = ["Continuar", "Salir", "Default"]

#Messages
msg = "Bienvenido. Elija su archivo de Excel con datos de PRIME"
title = "Bienvenido"

HIGH_PERCENTAGE = 1.1
LOW_PERCENTAGE = 0.9