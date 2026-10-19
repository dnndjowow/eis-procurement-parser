from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
EXCEL_DIR = BASE_DIR / "files_excel"

EXCEL_PATH = EXCEL_DIR / "Shapka_New.xlsx"
MAIN_PUBLICATION_EXCEL_PATH = EXCEL_DIR / "Publikatsia_zakupochnykh_protsedur_v_2026_godu.xlsx"
PUBLICATION_SYNC_LAST_COLUMN = "DH"
FILES_DIR = BASE_DIR / "files_kontract"
FAS_DIR = BASE_DIR / "files_fas"
REPORTS_DIR = BASE_DIR / "reports"
EIS_CA_BUNDLE = BASE_DIR / "certs" / "eis_cacert.pem"

HEADER_ROW = 1
NOTICE_COLUMN = "Основание для заключения Номер извещения/ распоряжение/ ст. № 44-ФЗ"
NOTICE_NUMBER_PATTERN = r"\b087340000392[678]\d{6}\b"
DATE_CELL_FORMAT = "dd.mm.yyyy"
TARGET_BUDGET_YEAR = 2026
FULL_CONTRACT_PRICE_FIELD = "Полная цена контракта"
FULL_EIS_CONTRACT_PRICE_FIELD = "Полная цена контракта ЕИС"

EIS_RETRY_ATTEMPTS = 3
EIS_RETRY_DELAY_SECONDS = 2
FAS_PRICE_THRESHOLD = 500_000_000

STAGE_NAMES = ["I", "II", "III", "IV", "V"]
STAGE_SUFFIXES = {
    "I": "I этапу",
    "II": "II этапу",
    "III": "III этапу",
    "IV": "IV этапу",
    "V": "V этапу",
}

COUNTRY_SUSPICIOUS_MARKERS = [
    "товар",
    "количество",
    "единица",
    "сертификат",
    "подтверждение",
    "итого",
]
