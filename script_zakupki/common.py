import re
from datetime import date, datetime

from openpyxl.utils.datetime import from_excel


MONTHS = {
    "января": 1,
    "февраля": 2,
    "марта": 3,
    "апреля": 4,
    "мая": 5,
    "июня": 6,
    "июля": 7,
    "августа": 8,
    "сентября": 9,
    "октября": 10,
    "ноября": 11,
    "декабря": 12,
}


def normalize_text(value):
    return " ".join(str(value or "").replace("\xa0", " ").replace("\u202f", " ").split())


def normalize_header(value):
    return normalize_text(value).lower()


def parse_number(value):
    text = normalize_text(value)
    if not text or text in {"-", "–"}:
        return None

    match = re.search(r"\d[\d\s\u202f]*(?:[,.]\d+)?", text)
    if not match:
        return None

    number = float(match.group(0).replace(" ", "").replace("\u202f", "").replace(",", "."))
    if number.is_integer():
        return int(number)

    return number


def parse_date(value):
    text = normalize_text(value).lower()

    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", text)
    if match:
        day, month, year = map(int, match.groups())
        return date(year, month, day)

    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", text)
    if match:
        year, month, day = map(int, match.groups())
        return date(year, month, day)

    match = re.search(
        r"[«\"]?(\d{1,2})[»\"]?\s+"
        r"(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
        r"\s+(\d{4})",
        text,
    )
    if match:
        day = int(match.group(1))
        month = MONTHS[match.group(2)]
        year = int(match.group(3))
        return date(year, month, day)

    return None


def parse_all_dates(value):
    text = normalize_text(value).lower()
    dates = []

    for match in re.finditer(r"(\d{2})\.(\d{2})\.(\d{4})", text):
        day, month, year = map(int, match.groups())
        dates.append(date(year, month, day))

    russian_date_pattern = (
        r"[«\"]?(\d{1,2})[»\"]?\s+"
        r"(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
        r"\s+(\d{4})"
    )
    for match in re.finditer(russian_date_pattern, text):
        day = int(match.group(1))
        month = MONTHS[match.group(2)]
        year = int(match.group(3))
        dates.append(date(year, month, day))

    return dates


def normalize_compare_value(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        parsed_date = parse_date(value)
        if parsed_date:
            return parsed_date
        return normalize_text(value)
    if isinstance(value, (int, float)):
        try:
            return from_excel(value).date()
        except Exception:
            return value
    return value


def values_equal(old_value, new_value):
    old_value = normalize_compare_value(old_value)
    new_value = normalize_compare_value(new_value)

    if old_value is None and new_value is None:
        return True
    if old_value is None or new_value is None:
        return False
    if isinstance(old_value, date) or isinstance(new_value, date):
        return old_value == new_value

    try:
        return abs(float(str(old_value).replace(",", ".")) - float(str(new_value).replace(",", "."))) < 0.01
    except ValueError:
        return str(old_value).strip() == str(new_value).strip()


def extract_notice_number(value):
    match = re.search(r"\b087\d{16}\b", str(value or ""))
    return match.group(0) if match else None


def extract_ikz(value):
    match = re.search(r"\b\d{36}\b", str(value or ""))
    return match.group(0) if match else None
