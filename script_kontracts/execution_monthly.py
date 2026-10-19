import copy
from dataclasses import dataclass

from excel_layout import (
    EXECUTION_COLUMN_MAP,
    EXECUTION_FIELD_ALIASES,
    header_matches,
)
from excel_tools import (
    CellChange,
    get_compare_cell_value,
    get_date_value,
    get_number_value,
    values_are_equal,
)
from journal_register import get_public_value
from settings import DATE_CELL_FORMAT, STAGE_NAMES, STAGE_SUFFIXES


PUBLIC_SOURCE_FIELDS = {
    "Нормативно-правовой акт": (
        "Нормативно-правовой акт, в рамках которого осуществляется централизованная закупка лекарственных препаратов"
    ),
    "МНН": "МНН закупаемого лекарственного препарата",
}

SMALL_EIS_PRICE_DIFF_RATIO = 0.01


@dataclass
class ExecutionMonthlyResult:
    changed: bool
    changes: list[CellChange]
    warnings: list[str]
    rows_written: int


def build_execution_columns(sheet):
    warnings = []

    for field, column in EXECUTION_COLUMN_MAP.items():
        header_value = sheet[f"{column}1"].value
        aliases = EXECUTION_FIELD_ALIASES[field]
        if not header_matches(header_value, aliases):
            warnings.append(f"Исп по мес: колонка не найдена: {field}")

    if warnings:
        warnings.insert(0, "Исп по мес: структура листа не соответствует Shapka_Example")
        return {}, warnings

    return dict(EXECUTION_COLUMN_MAP), warnings


def build_execution_rows(public_sheet, public_row, data, public_values_sheet=None):
    unit_price, amount_unit_price = get_unit_prices(public_sheet, public_row, data, public_values_sheet)
    execution_contract_price = get_execution_contract_price(data)
    base_row = {
        "Нормативно-правовой акт": get_public_field(
            public_sheet,
            public_row,
            "Нормативно-правовой акт",
            public_values_sheet,
        ),
        "Поставщик": data.get("Поставщик"),
        "Дата контракта": data.get("Дата заключения контракта"),
        "Номер контракта": data.get("Номер контракта"),
        "МНН": get_public_field(public_sheet, public_row, "МНН", public_values_sheet),
        "Цена за ед. товара, руб.": unit_price,
    }

    stage_rows = []
    has_stage_quantities = any(
        data.get(f"Количество товара по {STAGE_SUFFIXES[stage]}") not in (None, "")
        for stage in STAGE_NAMES
    )

    for stage in STAGE_NAMES:
        suffix = STAGE_SUFFIXES[stage]
        quantity = data.get(f"Количество товара по {suffix}")
        if stage == "I" and quantity in (None, "") and not has_stage_quantities:
            quantity = data.get("Общее количество товара")

        delivery_date = data.get(f"Срок поставки по {suffix}")
        document_date = data.get(f"Срок предоставления документов по {suffix}")

        if quantity in (None, "") and delivery_date in (None, "") and document_date in (None, ""):
            continue
        row = dict(base_row)
        row["_stage"] = stage
        row["_stage_title"] = f"{stage} этап"
        row["Сумма ГК/этапа к оплате, руб."] = calculate_stage_amount(quantity, amount_unit_price)
        row["Кол-во товара"] = quantity
        row["Срок поставки Товара"] = delivery_date
        row["Срок предоставления документов в ЕИС"] = document_date
        stage_rows.append(row)

    apply_contract_price_to_single_execution_row(stage_rows, execution_contract_price)
    apply_eis_price_to_execution_rows(stage_rows, execution_contract_price, amount_unit_price, data)
    return stage_rows


def get_execution_contract_price(data):
    eis_price = get_number_value(data, "Цена контракта ЕИС")
    if eis_price is not None:
        return eis_price

    return get_number_value(data, "Цена контракта")


def apply_eis_price_to_execution_rows(rows, execution_contract_price, unit_price, data):
    eis_price = get_number_value(data, "Цена контракта ЕИС")
    unit_price = get_number_value({"unit_price": unit_price}, "unit_price")

    if eis_price is None or execution_contract_price is None or not rows:
        return

    if len(rows) != 1:
        return

    current_quantity = get_number_value(rows[0], "Кол-во товара")
    word_price = get_number_value(data, "Цена контракта")
    if (
        current_quantity is not None
        and word_price is not None
        and word_price > 0
        and abs(eis_price - word_price) / word_price < SMALL_EIS_PRICE_DIFF_RATIO
    ):
        rows[0]["Цена за ед. товара, руб."] = normalize_number_for_excel(
            round(execution_contract_price / current_quantity, 2)
        )
        rows[0]["Сумма ГК/этапа к оплате, руб."] = normalize_number_for_excel(round(execution_contract_price, 2))
        return

    if unit_price is None:
        return

    quantity = execution_contract_price / unit_price
    rows[0]["Кол-во товара"] = normalize_number_for_excel(round(quantity, 4))
    rows[0]["Сумма ГК/этапа к оплате, руб."] = normalize_number_for_excel(round(execution_contract_price, 2))


def apply_contract_price_to_single_execution_row(rows, execution_contract_price):
    if len(rows) != 1 or execution_contract_price is None:
        return

    rows[0]["Сумма ГК/этапа к оплате, руб."] = normalize_number_for_excel(round(execution_contract_price, 2))


def calculate_stage_amount(quantity, unit_price):
    quantity_number = get_number_value({"quantity": quantity}, "quantity")
    unit_price_number = get_number_value({"unit_price": unit_price}, "unit_price")

    if quantity_number is None or unit_price_number is None:
        return None

    amount = float(round(quantity_number * unit_price_number, 2))
    if amount.is_integer():
        return int(amount)

    return amount


def normalize_number_for_excel(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)

    return value


def build_stage_amount_warnings(rows, data):
    warnings = []
    amounts = []

    for row in rows:
        stage_title = row.get("_stage_title", "этап")
        amount = get_number_value(row, "Сумма ГК/этапа к оплате, руб.")
        quantity = get_number_value(row, "Кол-во товара")
        unit_price = get_number_value(row, "Цена за ед. товара, руб.")

        if amount is None:
            missing = []
            if quantity is None:
                missing.append("количество")
            if unit_price is None:
                missing.append("цена за единицу")
            reason = " и ".join(missing) if missing else "данные"
            warnings.append(f"Исп по мес: не рассчитана сумма {stage_title}: нет {reason}")
            continue

        amounts.append(amount)

    contract_price = get_execution_contract_price(data)
    if contract_price is None or len(amounts) != len(rows):
        return warnings

    stages_total = round(sum(amounts), 2)
    if abs(stages_total - contract_price) > 1:
        warnings.append(
            "Исп по мес: сумма этапов "
            f"{format_number(stages_total)} не равна цене контракта {format_number(contract_price)}"
        )

    return warnings


def format_number(value):
    if value is None:
        return "пусто"

    if float(value).is_integer():
        return f"{int(value):,}".replace(",", " ")

    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def get_public_field(public_sheet, row, field, public_values_sheet=None):
    header = PUBLIC_SOURCE_FIELDS[field]
    try:
        return get_public_value(public_sheet, row, header, public_values_sheet)
    except ValueError:
        return None


def get_unit_prices(public_sheet, row, data, public_values_sheet=None):
    for header in ("Цена за ед. по ГК", "Цена за ед. по ТЗ"):
        try:
            value = get_public_value(public_sheet, row, header, public_values_sheet)
        except ValueError:
            continue

        if value not in (None, ""):
            unit_price = get_number_value({"Цена за ед.": value}, "Цена за ед.")
            if unit_price and unit_price > 0:
                return value, value

    contract_price = get_number_value(data, "Цена контракта")
    total_quantity = get_number_value(data, "Общее количество товара")
    if contract_price is not None and total_quantity:
        unit_price = round(contract_price / total_quantity, 2)
        return unit_price, unit_price

    return None, None


def find_execution_target_rows(sheet, columns, contract_number, rows_count):
    existing_rows = find_existing_contract_rows(sheet, columns, contract_number)
    target_rows = []
    start_row = max(existing_rows) + 1 if existing_rows else 2

    for index in range(rows_count):
        if index < len(existing_rows):
            target_rows.append(existing_rows[index])
            continue

        target_rows.append(find_first_empty_execution_row(
            sheet,
            columns,
            reserved_rows=target_rows,
            start_row=start_row,
        ))

    return target_rows, existing_rows


def find_existing_contract_rows(sheet, columns, contract_number):
    if not contract_number or "Номер контракта" not in columns:
        return []

    column = columns["Номер контракта"]
    contract_number = str(contract_number).strip()
    rows = []

    for row in range(2, sheet.max_row + 1):
        current_value = str(sheet[f"{column}{row}"].value or "").strip()
        if current_value == contract_number:
            rows.append(row)

    return rows


def find_first_empty_execution_row(sheet, columns, reserved_rows=None, start_row=2):
    reserved_rows = set(reserved_rows or [])
    check_columns = [
        columns[field]
        for field in ("Номер контракта", "Поставщик", "МНН")
        if field in columns
    ]
    if not check_columns:
        check_columns = list(columns.values())

    for row in range(start_row, sheet.max_row + 1):
        if row in reserved_rows:
            continue

        if all(sheet[f"{column}{row}"].value in (None, "") for column in check_columns):
            return row

    next_row = sheet.max_row + 1
    if reserved_rows:
        next_row = max(next_row, max(reserved_rows) + 1)

    return next_row


def ensure_row_style(sheet, row):
    if row <= sheet.max_row:
        return

    source_row = max(2, row - 1)
    for column in range(1, sheet.max_column + 1):
        source_cell = sheet.cell(source_row, column)
        target_cell = sheet.cell(row, column)
        if source_cell.has_style:
            target_cell._style = copy.copy(source_cell._style)
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format
        if source_cell.alignment:
            target_cell.alignment = copy.copy(source_cell.alignment)


def is_date_execution_field(field):
    return field.startswith("Дата") or field.startswith("Срок")


def is_money_execution_field(field):
    return field.startswith("Сумма") or field.startswith("Цена")


def write_execution_cell(sheet, values_sheet, row, column, field, value, changes, apply_changes):
    cell = sheet[f"{column}{row}"]
    old_value = get_compare_cell_value(cell, values_sheet)

    if values_are_equal(old_value, value, field):
        return False

    changes.append(CellChange(sheet.title, cell.coordinate, field, old_value, value))
    if apply_changes:
        cell.value = value
        if is_date_execution_field(field):
            cell.number_format = DATE_CELL_FORMAT
        elif is_money_execution_field(field):
            cell.number_format = "#,##0.00"

    return True


def clear_execution_cell(sheet, values_sheet, row, column, field, changes, apply_changes):
    cell = sheet[f"{column}{row}"]
    old_value = get_compare_cell_value(cell, values_sheet)
    if old_value in (None, ""):
        return False

    changes.append(CellChange(sheet.title, cell.coordinate, field, old_value, None))
    if apply_changes:
        cell.value = None

    return True


def write_execution_by_month(
    execution_sheet,
    public_sheet,
    row,
    data,
    public_values_sheet=None,
    execution_values_sheet=None,
    apply_changes=True,
):
    columns, warnings = build_execution_columns(execution_sheet)
    if warnings:
        return ExecutionMonthlyResult(False, [], warnings, 0)

    rows = build_execution_rows(public_sheet, row, data, public_values_sheet)
    if not rows:
        return ExecutionMonthlyResult(False, [], ["Исп по мес: не найдены этапы поставки"], 0)

    warnings.extend(build_stage_amount_warnings(rows, data))
    target_rows, existing_rows = find_execution_target_rows(
        execution_sheet,
        columns,
        data.get("Номер контракта"),
        len(rows),
    )

    changed = False
    changes = []

    for target_row, row_data in zip(target_rows, rows):
        ensure_row_style(execution_sheet, target_row)

        for field, value in row_data.items():
            if field.startswith("_"):
                continue
            if field not in columns or value in (None, ""):
                continue

            changed = write_execution_cell(
                execution_sheet,
                execution_values_sheet,
                target_row,
                columns[field],
                field,
                value,
                changes,
                apply_changes,
            ) or changed

    extra_rows = [
        existing_row
        for existing_row in existing_rows
        if existing_row not in target_rows
    ]
    for extra_row in extra_rows:
        for field, column in columns.items():
            changed = clear_execution_cell(
                execution_sheet,
                execution_values_sheet,
                extra_row,
                column,
                field,
                changes,
                apply_changes,
            ) or changed

    return ExecutionMonthlyResult(changed, changes, warnings, len(rows))
