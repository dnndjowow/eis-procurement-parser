import re

from config import field_mappings_public, optional_field_keys
from docx_parser.common import normalize_text
from document_search import append_unique
from excel_tools import get_date_value, get_number_value
from settings import COUNTRY_SUSPICIOUS_MARKERS, STAGE_NAMES, STAGE_SUFFIXES


class CheckInfo:
    def __init__(self, result_data):
        self.result_data = result_data

    def missing_fields(self):
        missing = []

        for field_key, _ in field_mappings_public:
            if field_key in optional_field_keys:
                continue

            if field_key == 'Производство готовой лекарственной формы/Страна происхождения Товара':
                path1 = field_key.split('/')[0]
                path2 = field_key.split('/')[1]
                value1 = self.result_data.get(path1)
                value2 = self.result_data.get(path2)
                if not is_filled_value(value1) and not is_filled_value(value2):
                    append_unique(missing, field_key)
                continue

            value = self.result_data.get(field_key)
            if not is_filled_value(value):
                append_unique(missing, field_key)

        return missing

    def check(self, verbose=True):
        missing = self.missing_fields()

        if not missing:
            if verbose:
                print('Все данные прошли проверку')
            return True

        if verbose:
            print(f'Ошибка проверки: не найдены поля: {", ".join(missing)}')

        return False


def is_filled_value(value):
    return value is not None and value != '' and not str(value).startswith('Ошибка')


def validate_data(data, verbose=False):
    checker = CheckInfo(data)
    return checker.check(verbose=verbose)


def get_missing_fields(data):
    checker = CheckInfo(data)
    return checker.missing_fields()


def count_excel_fields(data):
    filled = 0
    total = 0

    for field_key, _ in field_mappings_public:
        if '/' in field_key:
            field_parts = field_key.split('/')
            field_filled = any(is_filled_value(data.get(field_part)) for field_part in field_parts)
        else:
            field_filled = is_filled_value(data.get(field_key))

        if field_key in optional_field_keys:
            if field_filled:
                total += 1
                filled += 1
            continue

        total += 1
        if field_filled:
            filled += 1

    return filled, total


def check_suspicious_data(data):
    warnings = []
    warnings.extend(check_stage_quantity_sum(data))
    warnings.extend(check_stage_completeness(data))
    warnings.extend(check_document_deadlines(data))
    warnings.extend(check_stage_years(data))
    warnings.extend(check_positive_quantities(data))
    warnings.extend(check_positive_contract_price(data))
    warnings.extend(check_contract_price_eis_difference(data))
    warnings.extend(check_country_origin(data))
    return warnings


def check_stage_completeness(data):
    warnings = []

    for stage in STAGE_NAMES:
        suffix = STAGE_SUFFIXES[stage]
        quantity = get_number_value(data, f"Количество товара по {suffix}")
        delivery_date = get_date_value(data, f"Срок поставки по {suffix}")
        document_date = get_date_value(data, f"Срок предоставления документов по {suffix}")

        if quantity is not None and not delivery_date:
            warnings.append(f'по {suffix} есть количество, но не найден срок поставки')
        if quantity is not None and not document_date:
            warnings.append(f'по {suffix} есть количество, но не найден срок предоставления документов')
        if quantity is None and (delivery_date or document_date):
            warnings.append(f'по {suffix} есть срок, но не найдено количество товара')

    return warnings


def check_stage_quantity_sum(data):
    warnings = []
    stage_total = 0
    found_stage_quantity = False

    for stage in STAGE_NAMES:
        quantity = get_number_value(data, f"Количество товара по {STAGE_SUFFIXES[stage]}")
        if quantity is not None:
            found_stage_quantity = True
            stage_total += quantity

    total_quantity = get_number_value(data, "Общее количество товара")
    if found_stage_quantity and total_quantity is not None and abs(stage_total - total_quantity) > 0.01:
        warnings.append(
            f'сумма количеств по этапам {format_number(stage_total)} '
            f'не равна общему количеству {format_number(total_quantity)}'
        )

    return warnings


def check_document_deadlines(data):
    warnings = []

    for stage in STAGE_NAMES:
        suffix = STAGE_SUFFIXES[stage]
        delivery_date = get_date_value(data, f"Срок поставки по {suffix}")
        document_date = get_date_value(data, f"Срок предоставления документов по {suffix}")

        if delivery_date and document_date and document_date < delivery_date:
            warnings.append(
                f'срок предоставления документов по {suffix} '
                f'{document_date.strftime("%d.%m.%Y")} раньше срока поставки '
                f'{delivery_date.strftime("%d.%m.%Y")}'
            )

    return warnings


def check_stage_years(data):
    warnings = []
    period_years = normalize_years(data.get("Период поставок по годам"))
    if not period_years:
        return warnings

    for stage in STAGE_NAMES:
        suffix = STAGE_SUFFIXES[stage]
        for field_prefix in ("Срок поставки по", "Срок предоставления документов по"):
            field_key = f"{field_prefix} {suffix}"
            stage_date = get_date_value(data, field_key)
            if stage_date and stage_date.year not in period_years:
                warnings.append(
                    f'{field_key} относится к {stage_date.year}, '
                    f'а период поставки {format_year_period(period_years)}'
                )

    return warnings


def normalize_years(value):
    return sorted({int(year) for year in re.findall(r"\b20\d{2}\b", str(value or ""))})


def format_year_period(years):
    return "-".join(str(year) for year in sorted(set(years)))


def check_positive_quantities(data):
    warnings = []

    quantity_fields = [
        "Общее количество товара",
        *[
            f"Количество товара по {STAGE_SUFFIXES[stage]}"
            for stage in STAGE_NAMES
        ],
    ]

    for field_key in quantity_fields:
        quantity = get_number_value(data, field_key)
        if quantity is not None and quantity <= 0:
            warnings.append(f'{field_key} <= 0: {format_number(quantity)}')

    return warnings


def check_positive_contract_price(data):
    warnings = []
    price = get_number_value(data, "Цена контракта")

    if price is not None and price <= 0:
        warnings.append(f'цена контракта <= 0: {format_number(price)}')

    return warnings


def check_contract_price_eis_difference(data):
    warnings = []
    word_price = get_number_value(data, "Цена контракта")
    eis_price = get_number_value(data, "Цена контракта ЕИС")

    if word_price is not None and eis_price is not None:
        warnings.append(
            'цена контракта из Word отличается от ЕИС: '
            f'AE={format_number(word_price)}, AF={format_number(eis_price)}'
        )

    return warnings


def check_country_origin(data):
    warnings = []

    country = normalize_text(data.get("Страна происхождения Товара", ""))
    country_lower = country.lower()
    if country and any(marker in country_lower for marker in COUNTRY_SUSPICIOUS_MARKERS):
        warnings.append(f'страна происхождения выглядит подозрительно: {country}')

    return warnings


def format_number(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).replace('.', ',')
