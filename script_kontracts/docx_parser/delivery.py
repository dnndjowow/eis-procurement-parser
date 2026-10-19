import re
from datetime import date
from typing import Dict, List, Optional, Union

from docx import Document

from .common import collect_document_lines, normalize_text


STAGE_FIELD_SUFFIXES = {
    "I": "I этапу",
    "II": "II этапу",
    "III": "III этапу",
    "IV": "IV этапу",
    "V": "V этапу",
}

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


def parse_stage(value: str) -> Optional[str]:
    text = normalize_text(value).lower()
    exact_value = text.strip(" .)")
    standalone_stages = {
        "i": "I",
        "1": "I",
        "ii": "II",
        "2": "II",
        "iii": "III",
        "3": "III",
        "iv": "IV",
        "4": "IV",
        "v": "V",
        "5": "V",
    }
    if exact_value in standalone_stages:
        return standalone_stages[exact_value]

    match = re.search(
        r'\b(i{1,3}|iv|v|1|2|3|4|5|перв\w*|втор\w*|трет\w*|четвер\w*|пят\w*)\s+этап',
        text,
    )
    if not match:
        return None

    token = match.group(1)
    exact_stages = {
        "i": "I",
        "1": "I",
        "ii": "II",
        "2": "II",
        "iii": "III",
        "3": "III",
        "iv": "IV",
        "4": "IV",
        "v": "V",
        "5": "V",
    }

    if token in exact_stages:
        return exact_stages[token]

    russian_prefixes = [
        ("перв", "I"),
        ("втор", "II"),
        ("трет", "III"),
        ("четвер", "IV"),
        ("пят", "V"),
    ]
    for prefix, stage in russian_prefixes:
        if token.startswith(prefix):
            return stage

    return None


def parse_number(value: str) -> Optional[Union[int, float]]:
    text = normalize_text(value)
    if not text or text in {"-", "–"}:
        return None

    match = re.search(r'\d[\d\s\u202f]*(?:[,.]\d+)?', text)
    if not match:
        return None

    text = match.group(0).replace(" ", "").replace("\u202f", "").replace(",", ".")
    number = float(text)
    if number.is_integer():
        return int(number)

    return number


def normalize_key(value: str) -> str:
    return re.sub(r"[^а-яa-z0-9]", "", normalize_text(value).lower().replace("ё", "е"))


def parse_date(value: str) -> Optional[date]:
    text = normalize_text(value).lower()

    match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', text)
    if match:
        day, month, year = map(int, match.groups())
        return date(year, month, day)

    match = re.search(
        r'[«"]?(\d{1,2})[»"]?\s+'
        r'(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)'
        r'\s+(\d{4})',
        text
    )
    if match:
        day = int(match.group(1))
        month = MONTHS[match.group(2)]
        year = int(match.group(3))
        return date(year, month, day)

    return None


def parse_all_date_matches(value: str) -> List[tuple[date, int, int]]:
    text = normalize_text(value).lower()
    dates = []

    for match in re.finditer(r'(\d{2})\.(\d{2})\.(\d{4})', text):
        day, month, year = map(int, match.groups())
        dates.append((date(year, month, day), match.start(), match.end()))

    russian_date_pattern = (
        r'[«"]?(\d{1,2})[»"]?\s+'
        r'(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)'
        r'\s+(\d{4})'
    )
    for match in re.finditer(russian_date_pattern, text):
        day = int(match.group(1))
        month = MONTHS[match.group(2)]
        year = int(match.group(3))
        dates.append((date(year, month, day), match.start(), match.end()))

    return sorted(dates, key=lambda item: item[1])


def parse_all_dates(value: str) -> List[date]:
    dates = []

    for parsed_date, _, _ in parse_all_date_matches(value):
        dates.append(parsed_date)

    return dates


def extract_table_data(doc: Document) -> List[List[List[str]]]:
    tables = []

    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [normalize_text(cell.text) for cell in row.cells]
            if any(row_data):
                table_data.append(row_data)
        if table_data:
            tables.append(table_data)

    return tables


def parse_delivery_stage_table(doc: Document) -> Dict[str, object]:
    result = {}

    for table_data in extract_table_data(doc):
        header_index = -1
        for row_index, row in enumerate(table_data):
            row_text = ' '.join(row).lower()
            if is_delivery_header_text(row_text):
                header_index = row_index
                break

        if header_index == -1:
            continue

        header_rows = table_data[header_index:header_index + 3]
        stage_col, delivery_date_col, quantity_cols = detect_delivery_columns(header_rows)
        if stage_col is None or delivery_date_col is None:
            continue

        for row in table_data[header_index + 1:]:
            if len(row) <= max(stage_col, delivery_date_col):
                continue

            stage = parse_stage(row[stage_col])
            if not stage:
                continue

            delivery_dates = parse_all_dates(row[delivery_date_col])
            if delivery_dates:
                result[f"Срок поставки по {STAGE_FIELD_SUFFIXES[stage]}"] = delivery_dates[-1]

            quantity_cells = [
                row[column]
                for column in quantity_cols
                if len(row) > column
            ]
            if not quantity_cells:
                quantity_cells = row[delivery_date_col + 1:]

            quantities = [parse_number(cell) for cell in quantity_cells]
            quantities = [
                quantity
                for quantity in quantities
                if quantity is not None
            ]
            if quantities:
                result[f"Количество товара по {STAGE_FIELD_SUFFIXES[stage]}"] = sum(quantities)

    stage_quantity_fields = [
        f"Количество товара по {STAGE_FIELD_SUFFIXES[stage]}"
        for stage in STAGE_FIELD_SUFFIXES
    ]
    total_quantity = sum(
        result[field]
        for field in stage_quantity_fields
        if field in result
    )
    if total_quantity:
        result["Общее количество товара"] = total_quantity

    return result


def is_delivery_header_text(row_text: str) -> bool:
    row_key = normalize_key(row_text)
    has_stage = "этап" in row_key
    has_delivery = "поставк" in row_key or "отгрузк" in row_key
    has_deadline = "срок" in row_key or "дат" in row_key or "период" in row_key
    return has_stage and has_delivery and has_deadline


def detect_delivery_columns(header_rows):
    header_width = max(len(row) for row in header_rows)
    stage_col = None
    delivery_date_col = None
    quantity_cols = []

    for column_index in range(header_width):
        column_header = ' '.join(
            row[column_index].lower()
            for row in header_rows
            if len(row) > column_index
        )
        column_key = normalize_key(column_header)

        if stage_col is None and "этап" in column_key:
            stage_col = column_index
        if (
            delivery_date_col is None
            and ("срок" in column_key or "дат" in column_key or "период" in column_key)
            and ("поставк" in column_key or "отгрузк" in column_key)
        ):
            delivery_date_col = column_index
        if (
            "количеств" in column_key
            or "поставляем" in column_key
            or "объем" in column_key
            or "общее" in column_key
        ):
            quantity_cols.append(column_index)

    quantity_cols = [
        column
        for column in quantity_cols
        if column not in {stage_col, delivery_date_col}
    ]

    return stage_col, delivery_date_col, quantity_cols


def extract_clause_text(doc: Document, clause_number: str) -> str:
    lines = collect_document_lines(doc)
    clause_lines = []
    collecting = False
    clause_prefix = clause_number.rstrip(".")

    for line in lines:
        normalized_line = normalize_text(line)
        stripped_line = normalized_line.lstrip('«"')

        if not collecting:
            if stripped_line.startswith(clause_number):
                clause_lines.append(stripped_line)
                collecting = True
            continue

        clause_match = re.match(r'(\d+(?:\.\d+)+)\.', stripped_line)
        if clause_match:
            current_clause = clause_match.group(1)
            if current_clause != clause_prefix and not current_clause.startswith(f"{clause_prefix}."):
                break

        clause_lines.append(stripped_line)

    return " ".join(clause_lines)


def parse_document_deadlines(doc: Document) -> Dict[str, object]:
    result = {}
    clause = find_document_deadline_clause(doc)

    if not clause:
        return result

    stage_deadlines = parse_stage_deadlines(clause)
    if stage_deadlines:
        for stage, deadline in stage_deadlines.items():
            if stage in STAGE_FIELD_SUFFIXES:
                result[f"Срок предоставления документов по {STAGE_FIELD_SUFFIXES[stage]}"] = deadline
        return result

    dates = parse_all_dates(clause)
    if dates:
        result["Срок предоставления документов по I этапу"] = dates[-1]

    return result


def find_document_deadline_clause(doc: Document) -> str:
    for clause_number in ("9.4.", "9.3.", "9.5.", "10.1.", "10.2."):
        clause = extract_clause_text(doc, clause_number)
        if is_document_deadline_text(clause):
            return clause

    lines = collect_document_lines(doc)
    candidates = []
    for line in lines:
        line_lower = normalize_text(line).lower()
        if is_document_deadline_text(line_lower) and parse_all_dates(line_lower):
            candidates.append(line)

    return " ".join(candidates)


def is_document_deadline_text(text: str) -> bool:
    text = normalize_text(text).lower()
    if not text:
        return False

    has_document = "документ" in text or "еис" in text or "приемк" in text
    has_action = "предостав" in text or "размещ" in text or "формир" in text or "направ" in text
    return has_document and (has_action or bool(parse_all_dates(text)))


def parse_stage_deadlines(text: str) -> Dict[str, date]:
    deadlines = parse_date_before_stage_deadlines(text)
    stage_matches = list(re.finditer(
        r'(?:по|для|за)?\s*'
        r'(перв\w*|втор\w*|трет\w*|четвер\w*|пят\w*|i{1,3}|iv|v|1|2|3|4|5)\s+этап',
        text,
        flags=re.IGNORECASE
    ))

    for index, stage_match in enumerate(stage_matches):
        stage = parse_stage(stage_match.group(0))
        if not stage:
            continue
        if stage in deadlines:
            continue

        next_start = stage_matches[index + 1].start() if index + 1 < len(stage_matches) else len(text)
        after_segment = text[stage_match.end():next_start]
        dates = parse_all_dates(after_segment)
        if dates:
            deadlines[stage] = dates[0]

    return deadlines


def parse_date_before_stage_deadlines(text: str) -> Dict[str, date]:
    deadlines = {}

    for parsed_date, _, date_end in parse_all_date_matches(text):
        tail = text[date_end:date_end + 100]
        match = re.match(
            r'\s*(?:г\.)?\s*[-–—]\s*'
            r'(?:по|для|за)?\s*'
            r'(перв\w*|втор\w*|трет\w*|четвер\w*|пят\w*|i{1,3}|iv|v|1|2|3|4|5)\s+этап',
            normalize_text(tail).lower(),
            flags=re.IGNORECASE,
        )
        if not match:
            continue

        stage = parse_stage(match.group(0))
        if stage:
            deadlines[stage] = parsed_date

    return deadlines


def parse_delivery_terms(doc: Document) -> Dict[str, object]:
    result = {}
    result.update(parse_delivery_stage_table(doc))
    result.update(parse_document_deadlines(doc))
    return result
