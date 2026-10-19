from typing import List

from docx import Document


def normalize_text(text: str) -> str:
    """Нормализует пробелы, переносы и неразрывные пробелы."""
    return ' '.join(str(text or '').replace('\xa0', ' ').split())


def clean_price_value(price_text: str) -> str:
    return normalize_text(price_text).replace(' ', '')


def price_to_float(price_text: str) -> float:
    return float(clean_price_value(price_text).replace(',', '.'))


def collect_document_lines(doc: Document) -> List[str]:
    lines = []

    for paragraph in doc.paragraphs:
        text = normalize_text(paragraph.text)
        if text:
            lines.append(text)

    for table in doc.tables:
        for row in table.rows:
            row_text = ' | '.join(normalize_text(cell.text) for cell in row.cells)
            if row_text.strip(' |'):
                lines.append(row_text)

    return lines
