import argparse
import re
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    __package__ = "script_zakupki"

from .eis_client import fetch_purchase_data, find_notice_by_ikz
from .excel_writer import (
    clear_publication_data_rows,
    ensure_excel_is_closed,
    get_year_sheet,
    get_year_sheets,
    load_or_create_workbook,
    write_purchase_row,
)
from .logging_utils import TeeOutput, create_log_path
from .settings import BASE_DIR, EXCEL_PATH, TARGET_YEAR, ZAKUPKI_FILES_DIR
from .word_parser import iter_docx_files, parse_purchase_word


STAGE_TO_COLUMNS_BY_YEAR = {
    2026: {
        "I": ("AQ", "BA", "BE"),
        "II": ("AR", "BB", "BG"),
        "III": ("AS", "BC", "BI"),
        "IV": ("AT", "BD", "BK"),
    },
    2027: {
        "I": ("AQ", "BB", "BG"),
        "II": ("AR", "BC", "BI"),
        "III": ("AS", "BD", "BK"),
        "IV": ("AT", "BE", "BM"),
        "V": ("AU", "BF", "BO"),
    },
}

TOTAL_COLUMNS_BY_YEAR = {
    2026: {"children": "AU", "adults": "AW"},
    2027: {"children": "AV", "adults": "AX"},
}

STAGE_OUTPUT_COLUMNS_BY_YEAR = {
    2026: [
        "AQ", "AR", "AS", "AT",
        "AU", "AW",
        "BA", "BB", "BC", "BD",
        "BE", "BG", "BI", "BK",
    ],
    2027: [
        "AQ", "AR", "AS", "AT", "AU",
        "AV", "AX",
        "BB", "BC", "BD", "BE", "BF",
        "BG", "BI", "BK", "BM", "BO",
    ],
}


BASE_COLUMN_SHIFT_BY_YEAR = {
    2027: {
        "BM": "BQ",
        "BN": "BR",
        "BO": "BS",
        "BP": "BT",
        "BQ": "BU",
        "BR": "BV",
        "BS": "BW",
        "BT": "BX",
        "BU": "BY",
        "BV": "BZ",
        "BW": "CA",
        "BX": "CB",
        "BY": "CC",
        "BZ": "CD",
        "CA": "CE",
        "CB": "CF",
        "CC": "CG",
        "CD": "CH",
        "CE": "CI",
        "CF": "CJ",
        "CG": "CK",
        "CH": "CL",
        "CI": "CM",
        "CJ": "CN",
        "CK": "CO",
        "CL": "CP",
        "CM": "CQ",
        "CN": "CR",
        "CO": "CS",
        "CP": "CT",
        "CQ": "CU",
        "CR": "CV",
        "CS": "CW",
        "CT": "CX",
        "CU": "CY",
        "CV": "CZ",
        "CW": "DA",
        "CX": "DB",
        "CY": "DE",
        "CZ": "DF",
        "DA": "DG",
        "DB": "DH",
        "DC": "DI",
        "DD": "DJ",
        "DE": "DK",
        "DF": "DL",
        "DG": "DM",
        "DH": "DN",
    },
}


def collect_word_files(files_dir, word_paths):
    if word_paths:
        return [Path(path) for path in word_paths]

    return iter_docx_files(files_dir)


def build_purchase_rows(word_data):
    notice_number = word_data.notice_number
    if not notice_number:
        if not word_data.ikz:
            raise ValueError("в Word не найден ни номер извещения, ни ИКЗ")
        notice_number = find_notice_by_ikz(word_data.ikz)

    base_data = fetch_purchase_data(notice_number)

    if word_data.ikz:
        base_data["C"] = word_data.ikz
    if word_data.drug_name:
        base_data["V"] = word_data.drug_name

    budget_years = resolve_budget_years(word_data, base_data)
    budget_value = format_year_period(budget_years)
    delivery_years = resolve_delivery_years(word_data, budget_years)
    delivery_period = format_year_period(delivery_years) or base_data.get("B")

    rows = []
    for delivery_year in delivery_years:
        data = dict(base_data)
        apply_base_column_shift(data, delivery_year)
        if budget_value:
            data["A"] = budget_value
        if delivery_period:
            data["B"] = delivery_period
        data["_sheet_year"] = delivery_year
        data["_budget_year"] = resolve_financing_year_for_row(budget_years, delivery_year)
        data["_budget_years"] = budget_years
        data["_delivery_years"] = delivery_years
        data["_append_row"] = True
        add_stage_data(data, word_data, delivery_year)
        add_word_values(data, word_data, delivery_year)
        rows.append(data)

    return rows


def apply_base_column_shift(data, delivery_year):
    column_map = BASE_COLUMN_SHIFT_BY_YEAR.get(delivery_year, {})
    source_values = {
        source_column: data[source_column]
        for source_column in column_map
        if source_column in data
    }

    for source_column in source_values:
        data.pop(source_column, None)

    for source_column, target_column in column_map.items():
        if source_column not in source_values:
            continue

        value = source_values[source_column]
        if value in (None, ""):
            continue
        if data.get(target_column) in (None, ""):
            data[target_column] = value


def resolve_budget_years(word_data, data):
    if word_data.budget_years:
        return sorted(set(word_data.budget_years))

    financing_years = get_financing_years(data)
    if financing_years:
        return sorted(set(financing_years))

    return [TARGET_YEAR]


def resolve_delivery_years(word_data, fallback_years):
    years = sorted({
        stage.delivery_date.year
        for stage in word_data.stages
        if stage.delivery_date
    })
    if years:
        return years

    if word_data.delivery_stage_years:
        return list(word_data.delivery_stage_years)

    return list(fallback_years) or [TARGET_YEAR]


def resolve_financing_year_for_row(budget_years, delivery_year):
    if delivery_year in budget_years:
        return delivery_year
    if budget_years:
        return budget_years[0]
    return delivery_year


def format_year_period(years):
    if not years:
        return None

    return "-".join(str(year) for year in sorted(set(years)))


def add_stage_data(data, word_data, delivery_year):
    children_total = 0
    adult_total = 0
    generic_total = 0
    has_children = False
    has_adults = False
    has_generic = False

    children_formula_parts = []
    adult_formula_parts = []

    stage_columns = get_stage_columns(delivery_year)
    total_columns = get_total_columns(delivery_year)
    children_total_column = total_columns["children"]
    adult_total_column = total_columns["adults"]

    data["_clear_columns"] = get_stage_output_columns(delivery_year)

    target_year_stages = [
        stage
        for stage in word_data.stages
        if stage_matches_delivery_year(stage, delivery_year)
    ][:len(stage_columns)]
    data["_selected_stages"] = target_year_stages

    for index, stage_data in enumerate(target_year_stages):
        output_stage = list(stage_columns)[index]
        quantity_column, delivery_column, deadline_column = stage_columns[output_stage]

        if stage_data.total_quantity is not None:
            data[quantity_column] = make_stage_quantity_value(stage_data)
        if stage_data.delivery_date:
            data[delivery_column] = stage_data.delivery_date

        deadline = word_data.document_deadlines.get(stage_data.stage)
        if deadline:
            data[deadline_column] = deadline

        if stage_data.children_quantity is not None:
            children_total += stage_data.children_quantity
            has_children = True
            if stage_data.children_quantity:
                children_formula_parts.append(stage_data.children_quantity)
        if stage_data.adult_quantity is not None:
            adult_total += stage_data.adult_quantity
            has_adults = True
            if stage_data.adult_quantity:
                adult_formula_parts.append(stage_data.adult_quantity)
        if (
            stage_data.quantity is not None
            and stage_data.children_quantity is None
            and stage_data.adult_quantity is None
        ):
            generic_total += stage_data.quantity
            has_generic = True

    if has_children:
        data[children_total_column] = make_sum_value(children_formula_parts, fallback=children_total)
    elif has_adults:
        data[children_total_column] = 0

    if has_adults:
        data[adult_total_column] = make_sum_value(adult_formula_parts, fallback=adult_total)

    if has_generic and not has_children and not has_adults:
        if data.get("AI") and not data.get("AH"):
            data[children_total_column] = "=AP{row}"
        elif data.get("AH") and not data.get("AI"):
            data[children_total_column] = 0
            data[adult_total_column] = "=AP{row}"


def add_word_values(data, word_data, delivery_year):
    manual_formula_columns = []

    if word_data.package_quantity is not None:
        package_column = get_package_quantity_column(delivery_year)
        data[package_column] = word_data.package_quantity
        manual_formula_columns.append(package_column)

    if is_multiyear_purchase(data) and word_data.contract_unit_price is not None:
        data["AN"] = word_data.contract_unit_price
        manual_formula_columns.append("AN")

    if manual_formula_columns:
        data["_manual_formula_columns"] = manual_formula_columns


def get_package_quantity_column(delivery_year):
    if delivery_year == 2027:
        return "AZ"

    return "AY"


def get_stage_columns(delivery_year):
    return STAGE_TO_COLUMNS_BY_YEAR.get(delivery_year, STAGE_TO_COLUMNS_BY_YEAR[2026])


def get_total_columns(delivery_year):
    return TOTAL_COLUMNS_BY_YEAR.get(delivery_year, TOTAL_COLUMNS_BY_YEAR[2026])


def get_stage_output_columns(delivery_year):
    return STAGE_OUTPUT_COLUMNS_BY_YEAR.get(delivery_year, STAGE_OUTPUT_COLUMNS_BY_YEAR[2026])


def stage_matches_delivery_year(stage_data, delivery_year):
    if stage_data.delivery_date:
        return stage_data.delivery_date.year == delivery_year

    return False


def make_stage_quantity_value(stage_data):
    if stage_data.children_quantity is None and stage_data.adult_quantity is None:
        return stage_data.total_quantity

    parts = [
        value
        for value in (stage_data.children_quantity, stage_data.adult_quantity)
        if value
    ]
    return make_sum_value(parts, fallback=stage_data.total_quantity)


def make_sum_value(parts, fallback=0):
    if not parts:
        return fallback
    if len(parts) == 1:
        return parts[0]
    return "=" + "+".join(format_formula_number(part) for part in parts)


def format_formula_number(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace(",", ".")


def run_main(workbook_path=EXCEL_PATH, files_dir=ZAKUPKI_FILES_DIR, word_paths=None, write_enabled=False):
    if write_enabled:
        try:
            ensure_excel_is_closed(workbook_path)
        except RuntimeError as error:
            print(f"ERR: {error}")
            return

    word_files = collect_word_files(files_dir, word_paths)
    if not word_files:
        print(f"ERR: Word-файлы закупок не найдены: {files_dir}")
        return

    print("Режим проверки: запись в Excel отключена" if not write_enabled else "Режим записи в Excel")
    print(f"Excel: {workbook_path}")
    print(f"Word-файлов: {len(word_files)}")

    workbook = load_or_create_workbook(workbook_path)
    year_sheets = get_year_sheets(workbook)
    cleared_rows = clear_year_sheets(year_sheets)
    clear_action = "Excel очищен" if write_enabled else "Excel будет очищен"
    print(f"{clear_action}: старых строк {cleared_rows}")

    processed = 0
    changed = 0
    errors = 0
    warnings_count = 0
    problem_report_lines = []

    for word_file in word_files:
        try:
            word_data = parse_purchase_word(word_file)
            data_rows = build_purchase_rows(word_data)
        except Exception as error:
            errors += 1
            problem_report_lines.append(f"{word_file.name}:")
            problem_report_lines.append(f"    - ERR: {error}")
            print(f"ERR {word_file.name}: {error}")
            continue

        processed += 1

        for data in data_rows:
            sheet_year = data.get("_sheet_year")
            sheet = get_year_sheet(workbook, sheet_year)
            if sheet is None:
                warnings_count += 1
                identifier = data.get("D", word_file.name)
                warning = f"нет листа для года поставки {sheet_year}"
                problem_report_lines.append(f"{identifier}:")
                problem_report_lines.append(f"    - WARN: {warning}")
                print(f"WARN {identifier}: {warning}")
                continue

            item_warnings = check_purchase_warnings(data, word_data)
            if item_warnings:
                identifier = data.get("D", word_file.name)
                problem_report_lines.append(f"{identifier}:")
                for warning in item_warnings:
                    warnings_count += 1
                    problem_report_lines.append(f"    - WARN: {warning}")
                    print(f"WARN {identifier}: {warning}")

            result = write_purchase_row(sheet, data, apply_changes=True)
            if result.changed:
                changed += 1
                action = "записано" if write_enabled else "будет записано"
            else:
                action = "без изменений"

            print(
                f"OK  {data['D']}: лист {sheet.title}, строка {result.row}; "
                f"бюджет {data.get('A')}; период {data.get('B')}; "
                f"Word {word_file.name}; ячеек {len(result.changes)}; {action}"
            )

    if (changed or cleared_rows) and write_enabled:
        prepare_workbook_for_recalculation(workbook)
        workbook.save(workbook_path)

    print("")
    print("Итог:")
    print(f"Обработано Word-файлов: {processed}")
    print(f"Изменено строк: {changed}")
    print(f"Предупреждений: {warnings_count}")
    print(f"Ошибок: {errors}")
    print("")
    print("Проблемные закупки:")
    if problem_report_lines:
        for line in problem_report_lines:
            print(line)
    else:
        print("Проблемных закупок нет")


def clear_year_sheets(year_sheets):
    return sum(clear_publication_data_rows(sheet) for sheet in year_sheets.values())


def check_purchase_warnings(data, word_data):
    warnings = []
    sheet_year = data.get("_sheet_year") or TARGET_YEAR
    budget_label = data.get("A") or data.get("_budget_year") or TARGET_YEAR
    financing_year = sheet_year
    selected_stages = data.get("_selected_stages", [])

    if not selected_stages:
        if not word_data.delivery_stage_count:
            warnings.append("в Word не найдены этапы поставки")
        else:
            years = ", ".join(str(year) for year in word_data.delivery_stage_years) or "год не определен"
            warnings.append(f"нет этапов поставки для листа {sheet_year} года; в Word найдены годы: {years}")

    for stage in selected_stages:
        if stage.total_quantity is None:
            warnings.append(f"по {stage.stage} этапу не найдено количество товара")
        if not stage.delivery_date:
            warnings.append(f"по {stage.stage} этапу не найден срок поставки")

    if not data.get("AH") and not data.get("AI"):
        warnings.append(f"не найден КБК для бюджета {budget_label}")

    nmc_column = get_target_year_nmc_column(data, financing_year)
    if is_multiyear_purchase(data) and nmc_column != "W" and not data.get(nmc_column):
        warnings.append(f"многолетняя закупка: не найдена НМЦК {financing_year} года в колонке {nmc_column}")

    return warnings


def is_multiyear_purchase(data):
    years = get_financing_years(data)
    if len(years) > 1:
        return True

    return "-" in str(data.get("B") or "")


def get_target_year_nmc_column(data, target_year=None):
    target_year = target_year or TARGET_YEAR
    years = get_financing_years(data)
    year_columns = ["X", "Y", "Z"]

    if target_year in years:
        index = years.index(target_year)
        if index < len(year_columns):
            return year_columns[index]

    return "W"


def get_financing_years(data):
    years = data.get("_financing_years")
    if years:
        return [int(year) for year in years]

    return [int(year) for year in re.findall(r"20\d{2}", str(data.get("B") or ""))]


def main(workbook_path=EXCEL_PATH, files_dir=ZAKUPKI_FILES_DIR, word_paths=None, write_enabled=False):
    log_path = create_log_path()
    original_stdout = sys.stdout

    with open(log_path, "w", encoding="utf-8") as log_file:
        sys.stdout = TeeOutput(original_stdout, log_file)
        try:
            run_main(
                workbook_path=workbook_path,
                files_dir=files_dir,
                word_paths=word_paths,
                write_enabled=write_enabled,
            )
            print(f"Лог сохранен: {log_path.relative_to(BASE_DIR)}")
        finally:
            sys.stdout = original_stdout


def prepare_workbook_for_recalculation(workbook):
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True


def parse_args():
    parser = argparse.ArgumentParser(description="Парсер закупок")
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--check",
        action="store_true",
        help="проверить закупки без записи",
    )
    mode_group.add_argument(
        "--write",
        action="store_true",
        help="записать данные по закупкам в Excel",
    )
    parser.add_argument(
        "--excel",
        default=str(EXCEL_PATH),
        help="путь к Excel-файлу закупок",
    )
    parser.add_argument(
        "--files",
        default=str(ZAKUPKI_FILES_DIR),
        help="папка с Word-файлами закупок",
    )
    parser.add_argument(
        "--word",
        action="append",
        help="конкретный Word-файл закупки; можно указать несколько раз",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    main(
        workbook_path=Path(args.excel),
        files_dir=Path(args.files),
        word_paths=args.word,
        write_enabled=args.write,
    )
