from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
EXCEL_DIR = BASE_DIR / "files_excel"

EXCEL_PATH = EXCEL_DIR / "for_zakupki.xlsx"
ZAKUPKI_FILES_DIR = BASE_DIR / "files_zakupki"
REPORTS_DIR = BASE_DIR / "reports"
EIS_CA_BUNDLE = BASE_DIR / "certs" / "eis_cacert.pem"
TARGET_YEAR = 2026

DATE_FORMAT = "dd\\.mm\\.yyyy"

EIS_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
}

EIS_RETRY_ATTEMPTS = 3
EIS_RETRY_DELAY_SECONDS = 2
