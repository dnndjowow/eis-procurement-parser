import copy
import re
from dataclasses import dataclass

from excel_tools import (
    build_header_column_map,
    CellChange,
    get_compare_cell_value,
    normalize_header,
    resolve_column,
    values_are_equal,
)
from excel_layout import JOURNAL_FIELD_MAPPINGS, JOURNAL_SHEET_NAME
from settings import FULL_CONTRACT_PRICE_FIELD


PUBLIC_SOURCE_FIELDS = {
    "МНН закупаемого лекарственного препарата",
    "Нормативно-правовой акт, в рамках которого осуществляется централизованная закупка лекарственных препаратов",
}


@dataclass
class JournalRegisterResult:
    changed: bool
    changes: list[CellChange]
    warnings: list[str]
    row: int


def normalize_supplier_name(value):
    text = " ".join(str(value or "").split())
    if not text:
        return text

    if '"' in text:
        return text

    match = re.match(r"^(ООО|АО|ЗАО|ПАО|ИП)\s+(.+)$", text, flags=re.IGNORECASE)
    if not match:
        return text

    company_type = match.group(1).upper()
    name = match.group(2).strip()
    name = name.title()
    return f'{company_type} "{name}"'


def normalize_digits(value):
    return re.sub(r"\D", "", str(value or ""))


def split_codes(value):
    if value in (None, ""):
        return []

    text = str(value)
    return [
        part.strip()
        for part in re.split(r"[\n;,]+", text)
        if part.strip()
    ]


def get_kbk_target_article(kbk_code):
    digits = normalize_digits(kbk_code)
    if len(digits) >= 20:
        return digits[7:17]

    return ""


def get_kbk_entries(public_sheet, row, public_values_sheet=None):
    entries = []
    seen_codes = set()

    for header in get_kbk_headers(public_sheet):
        value = get_public_value(public_sheet, row, header, public_values_sheet)
        for code in split_codes(value):
            digits = normalize_digits(code)
            if not digits or digits in seen_codes:
                continue
            seen_codes.add(digits)
            entries.append({
                "КБК": digits,
                "Целевая статья": get_kbk_target_article(digits),
            })

    return entries


def get_kbk_headers(public_sheet):
    header_columns = build_header_column_map(public_sheet)
    sheet_year = get_sheet_year(public_sheet)
    preferred_headers = []

    if sheet_year:
        preferred_headers.extend([
            f"КБК {sheet_year} (основной)",
            f"КБК {sheet_year} (дети)",
        ])

    headers = []
    for header in preferred_headers:
        if normalize_header(header) in header_columns:
            headers.append(header)

    if headers:
        return headers

    for normalized_header in header_columns:
        if "кбк" in normalized_header and (
            "основной" in normalized_header
            or "дети" in normalized_header
        ):
            headers.append(normalized_header)

    return headers


def get_sheet_year(sheet):
    match = re.search(r"\b(20\d{2})\b", str(sheet.title))
    if match:
        return match.group(1)

    return ""


def get_kosgu_by_target_article(data):
    result = {}
    for row in data.get("Финансирование КОСГУ", []) or []:
        target_article = normalize_digits(row.get("Целевая статья"))
        kosgu_code = normalize_digits(row.get("КОСГУ"))
        if target_article and len(kosgu_code) == 3:
            result[target_article] = kosgu_code

    return result


def get_kosgu_codes(data):
    return [
        code
        for code in (normalize_digits(value) for value in data.get("Коды КОСГУ", []) or [])
        if len(code) == 3
    ]


def choose_kosgu_for_kbk(entry, entries, data, index):
    kosgu_by_target = get_kosgu_by_target_article(data)
    target_article = entry.get("Целевая статья")
    if target_article and target_article in kosgu_by_target:
        return kosgu_by_target[target_article]

    kosgu_codes = get_kosgu_codes(data)
    unique_kosgu_codes = []
    for code in kosgu_codes:
        if code not in unique_kosgu_codes:
            unique_kosgu_codes.append(code)

    if len(unique_kosgu_codes) == 1:
        return unique_kosgu_codes[0]

    if len(kosgu_codes) == len(entries):
        return kosgu_codes[index]

    return ""


def build_kosgu_value(public_sheet, row, data, public_values_sheet=None):
    entries = get_kbk_entries(public_sheet, row, public_values_sheet)
    formatted_codes = []
    warnings = []

    if not entries:
        return "", [f"нет КБК в листе {public_sheet.title}"]

    for index, entry in enumerate(entries):
        kosgu_code = choose_kosgu_for_kbk(entry, entries, data, index)
        kbk_code = entry["КБК"]

        if len(kbk_code) != 20:
            warnings.append(f"КБК {kbk_code} имеет длину {len(kbk_code)}, ожидается 20 цифр")

        if not kosgu_code:
            warnings.append(f"для КБК {kbk_code} не найден КОСГУ в пункте 9.1")
            formatted_codes.append(kbk_code)
            continue

        full_code = f"{kbk_code}{kosgu_code}"
        if len(full_code) != 23 or not full_code.isdigit():
            warnings.append(f"итоговый КОСГУ {full_code} имеет неверный формат, ожидается 23 цифры")

        formatted_codes.append(full_code)

    return "\n".join(formatted_codes), warnings


def format_kosgu(public_sheet, row, data, public_values_sheet=None, notice_number=None):
    value, warnings = build_kosgu_value(public_sheet, row, data, public_values_sheet)
    for warning in warnings:
        if notice_number:
            print(f"WARN {notice_number}: {warning}")

    return value


def normalize_multiline_value(value):
    return "\n".join(
        line.strip()
        for line in str(value or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if line.strip()
    )


def validate_written_kosgu(journal_sheet, target_row, expected_value):
    warnings = []
    column = resolve_column(journal_sheet, "КОСГУ")
    actual_value = journal_sheet[f"{column}{target_row}"].value
    actual_value = normalize_multiline_value(actual_value)
    expected_value = normalize_multiline_value(expected_value)

    if expected_value and actual_value != expected_value:
        warnings.append(f"КОСГУ в Журнал регистр не совпадает: ожидалось {expected_value}, записано {actual_value or 'пусто'}")

    return warnings


def journal_values_are_equal(field_key, old_value, new_value):
    if field_key == "КОСГУ":
        return normalize_multiline_value(old_value) == normalize_multiline_value(new_value)

    return values_are_equal(old_value, new_value, field_key)


def get_public_value(public_sheet, row, header, public_values_sheet=None):
    column = resolve_column(public_sheet, header)
    cell = public_sheet[f"{column}{row}"]
    if public_values_sheet is not None:
        value = public_values_sheet[cell.coordinate].value
        if value is not None:
            return value

    return cell.value


def build_journal_data(public_sheet, row, data, public_values_sheet=None, notice_number=None, warnings=None):
    journal_data = {}

    for field_key, _ in JOURNAL_FIELD_MAPPINGS:
        if field_key == "Номер извещения":
            journal_data[field_key] = notice_number
        elif field_key == "Поставщик":
            journal_data[field_key] = normalize_supplier_name(data.get(field_key))
        elif field_key == "КОСГУ":
            kosgu_value, kosgu_warnings = build_kosgu_value(public_sheet, row, data, public_values_sheet)
            journal_data[field_key] = kosgu_value
            if warnings is not None:
                warnings.extend(kosgu_warnings)
        elif field_key == "Цена контракта":
            journal_data[field_key] = data.get(FULL_CONTRACT_PRICE_FIELD) or data.get(field_key)
        elif field_key in PUBLIC_SOURCE_FIELDS:
            journal_data[field_key] = get_public_value(public_sheet, row, field_key, public_values_sheet)
        else:
            journal_data[field_key] = data.get(field_key)

    return journal_data


def is_positive_price(value):
    if value in (None, ""):
        return False

    try:
        number = float(str(value).replace(" ", "").replace("\u202f", "").replace(",", "."))
    except ValueError:
        return False

    return number > 0


def validate_journal_data(journal_data):
    warnings = []

    notice_number = normalize_digits(journal_data.get("Номер извещения"))
    if not notice_number:
        warnings.append("Журнал регистр: не найден номер извещения")

    supplier = str(journal_data.get("Поставщик") or "").strip()
    if not supplier:
        warnings.append("Журнал регистр: не найден поставщик")

    inn = normalize_digits(journal_data.get("ИНН"))
    if not inn:
        warnings.append("Журнал регистр: не найден ИНН поставщика")
    elif len(inn) not in (10, 12):
        warnings.append(f"Журнал регистр: ИНН {inn} имеет длину {len(inn)}, ожидается 10 или 12 цифр")

    if not is_positive_price(journal_data.get("Цена контракта")):
        warnings.append("Журнал регистр: не найдена или некорректна цена контракта")

    kosgu_value = normalize_multiline_value(journal_data.get("КОСГУ"))
    if not kosgu_value:
        warnings.append("Журнал регистр: не сформирован КОСГУ")
    else:
        for index, line in enumerate(kosgu_value.split("\n"), 1):
            code = normalize_digits(line)
            if code != line or len(code) != 23:
                warnings.append(
                    f"Журнал регистр: КОСГУ строка {index} имеет неверный формат "
                    f"({line}), ожидается 23 цифры"
                )

    return warnings


def find_journal_row(journal_sheet, notice_number, registry_number=None):
    search_values = [
        str(value).strip()
        for value in (notice_number, registry_number)
        if value not in (None, "")
    ]

    if not search_values:
        return find_first_empty_row(journal_sheet)

    column = resolve_column(journal_sheet, "основание для заключения")

    for row in range(2, journal_sheet.max_row + 1):
        current_value = str(journal_sheet[f"{column}{row}"].value or "").strip()
        if current_value in search_values:
            return row

    return find_first_empty_row(journal_sheet)


def find_first_empty_row(sheet):
    column = resolve_column(sheet, "основание для заключения")

    for row in range(2, sheet.max_row + 1):
        if sheet[f"{column}{row}"].value in (None, ""):
            return row

    row = sheet.max_row + 1
    copy_row_style(sheet, row - 1, row)
    return row


def copy_row_style(sheet, source_row, target_row):
    for column in range(1, sheet.max_column + 1):
        source_cell = sheet.cell(source_row, column)
        target_cell = sheet.cell(target_row, column)
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format
        if source_cell.alignment:
            target_cell.alignment = copy.copy(source_cell.alignment)


def write_journal_register(
    public_sheet,
    journal_sheet,
    row,
    data,
    notice_number=None,
    journal_values_sheet=None,
    public_values_sheet=None,
    apply_changes=True,
):
    warnings = []
    journal_data = build_journal_data(
        public_sheet,
        row,
        data,
        public_values_sheet,
        notice_number,
        warnings,
    )
    warnings.extend(validate_journal_data(journal_data))
    notice_value = journal_data.get("Номер извещения")
    registry_number = data.get("Номер реестровой записи контракта")
    target_row = find_journal_row(journal_sheet, notice_value, registry_number)
    header_columns = build_header_column_map(journal_sheet)
    changed = False
    changes = []

    for field_key, column_or_header in JOURNAL_FIELD_MAPPINGS:
        value = journal_data.get(field_key)
        if value in (None, ""):
            continue

        column = resolve_column(journal_sheet, column_or_header, header_columns)
        cell = journal_sheet[f"{column}{target_row}"]
        old_value = get_compare_cell_value(cell, journal_values_sheet)

        if journal_values_are_equal(field_key, old_value, value):
            continue

        changes.append(CellChange(
            JOURNAL_SHEET_NAME,
            cell.coordinate,
            field_key,
            old_value,
            value,
        ))
        if apply_changes:
            cell.value = value
        changed = True

    if apply_changes:
        warnings.extend(validate_written_kosgu(journal_sheet, target_row, journal_data.get("КОСГУ")))

    return JournalRegisterResult(changed=changed, changes=changes, warnings=warnings, row=target_row)
