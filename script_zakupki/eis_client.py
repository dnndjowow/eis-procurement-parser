import re
import time
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

from .common import normalize_text, parse_date, parse_number
from .settings import EIS_CA_BUNDLE, EIS_HEADERS, EIS_RETRY_ATTEMPTS, EIS_RETRY_DELAY_SECONDS


BASE_URL = "https://zakupki.gov.ru"

ADULT_KBK_CODES = {
    "05609010141892007244",
    "05609020141892007244",
    "05609020141892034244",
    "05609020141892089244",
    "05609020141892005244",
}

CHILD_KBK_CODES = {
    "05609020130592096244",
    "05609020130592099244",
}


def find_notice_by_ikz(ikz):
    params = {"searchString": ikz}
    url = f"{BASE_URL}/epz/order/extendedsearch/results.html?{urlencode(params)}"
    response = fetch_response(url)

    notices = []
    for match in re.finditer(r"\b087\d{16}\b", response.text):
        notice = match.group(0)
        if notice not in notices:
            notices.append(notice)

    if not notices:
        raise ValueError(f"ЕИС: не найден номер извещения по ИКЗ {ikz}")
    if len(notices) > 1:
        raise ValueError(f"ЕИС: по ИКЗ {ikz} найдено несколько извещений: {', '.join(notices)}")

    return notices[0]


def fetch_purchase_data(notice_number):
    common_soup = fetch_soup(
        f"{BASE_URL}/epz/order/notice/ea20/view/common-info.html?regNumber={notice_number}"
    )

    data = {}
    data["D"] = notice_number
    data["R"] = f"{BASE_URL}/epz/order/notice/ea20/view/common-info.html?regNumber={notice_number}"
    data.update(parse_common_info(common_soup))
    return data


def fetch_soup(url):
    response = fetch_response(url)
    response.encoding = "utf-8"
    return BeautifulSoup(response.text, "html.parser")


def fetch_response(url, attempts=EIS_RETRY_ATTEMPTS):
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            response = requests.get(
                url,
                headers=EIS_HEADERS,
                timeout=30,
                verify=get_eis_verify(),
            )
            response.raise_for_status()
            return response
        except Exception as error:
            last_error = error

        if attempt < attempts:
            print(f"WARN ЕИС: повтор {attempt}/{attempts} ({format_eis_error(last_error)})")
            time.sleep(EIS_RETRY_DELAY_SECONDS)

    raise last_error


def get_eis_verify():
    if EIS_CA_BUNDLE.exists():
        return str(EIS_CA_BUNDLE)

    return True


def format_eis_error(error):
    if isinstance(error, requests.exceptions.SSLError):
        return "ошибка SSL-сертификата ЕИС"
    if isinstance(error, requests.exceptions.Timeout):
        return "таймаут подключения"
    if isinstance(error, requests.exceptions.ConnectionError):
        return "ошибка подключения"
    if isinstance(error, requests.exceptions.HTTPError):
        response = error.response
        if response is not None:
            return f"HTTP {response.status_code}"
        return "HTTP-ошибка"

    return str(error)


def parse_common_info(soup):
    data = {}

    object_name = get_section_value(soup, "Наименование объекта закупки") or get_card_value(soup, "Объект закупки")
    data["V"] = parse_drug_from_object_name(object_name)
    data["P"] = parse_normative_act(object_name)
    data["E"] = parse_date(get_section_value(soup, "Дата и время начала срока подачи заявок"))
    data["G"] = parse_date(get_section_value(soup, "Дата и время окончания срока подачи заявок"))
    data["H"] = parse_date(get_section_value(soup, "Дата подведения итогов определения поставщика (подрядчика, исполнителя)"))
    data["W"] = parse_number(get_section_value(soup, "Начальная (максимальная) цена контракта"))
    data["C"] = normalize_text(get_section_value(soup, "Идентификационный код закупки (ИКЗ)"))
    data["AK"] = parse_date(get_section_value(soup, "Срок исполнения контракта"))
    data["BW"] = parse_security_percent(get_section_value(soup, "Размер обеспечения исполнения контракта"))
    data.update(parse_financing(soup))
    data.update(parse_kbk(soup))
    data.update(parse_medicine_table(soup))

    data["AE"] = 0
    data["CK"] = 0
    return {key: value for key, value in data.items() if value not in (None, "")}


def get_section_value(soup, label):
    label_normalized = normalize_text(label).lower()
    for title in soup.find_all("span", class_="section__title"):
        title_text = normalize_text(title.get_text(" ")).lower()
        if title_text == label_normalized:
            section = title.find_parent("section")
            if not section:
                continue
            info = section.find("span", class_="section__info")
            if info:
                return normalize_text(info.get_text(" "))
    return None


def get_card_value(soup, label):
    lines = [
        normalize_text(line)
        for line in soup.get_text("\n").splitlines()
        if normalize_text(line)
    ]
    for index, line in enumerate(lines):
        if line == label and index + 1 < len(lines):
            return lines[index + 1]
    return None


def parse_drug_from_object_name(value):
    text = normalize_text(value)
    if not text:
        return None

    match = re.search(
        r"Поставка иммунобиологического лекарственного препарата\s+для иммунопрофилактики\s+(.+?)\s+в рамках",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        return normalize_drug_name(match.group(1))

    match = re.search(
        r"Поставка иммунобиологического лекарственного препарата\s+(.+?)\s+в рамках",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        return normalize_drug_name(match.group(1))

    match = re.search(
        r"Поставка лекарственного препарата\s+(.+?)\s+в рамках",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        return normalize_drug_name(match.group(1))

    return text


def normalize_drug_name(value):
    text = normalize_text(value).strip(" .")
    return text[:1].upper() + text[1:] if text else None


def parse_normative_act(value):
    text = normalize_text(value)
    if not text:
        return None

    match = re.search(r"№\s*([0-9]+(?:\s*[А-ЯA-Z]+)?)", text)
    if match:
        act = normalize_text(match.group(1))
        return act

    return None


def parse_security_percent(value):
    match = re.search(r"\((\d+(?:[,.]\d+)?)\s*%\)", normalize_text(value))
    if not match:
        return None
    return parse_number(match.group(1))


def parse_financing(soup):
    result = {"X": 0, "Y": 0, "Z": 0}
    table = find_financing_table(soup)
    if not table:
        return result

    headers = [normalize_text(cell.get_text(" ")) for cell in table.find_all("tr")[0].find_all(["th", "td"])]
    rows = table.find_all("tr")
    if len(rows) < 2:
        return result

    values = [normalize_text(cell.get_text(" ")) for cell in rows[1].find_all(["th", "td"])]
    yearly_amounts = []

    for header, value in zip(headers, values):
        year_match = re.search(r"На\s+(\d{4})\s+год", header)
        if not year_match:
            continue
        amount = parse_number(value) or 0
        if amount:
            yearly_amounts.append((int(year_match.group(1)), amount))

    if yearly_amounts:
        result["B"] = "-".join(str(year) for year, _ in yearly_amounts)
        result["_financing_years"] = [year for year, _ in yearly_amounts]

    if len(yearly_amounts) > 1:
        for column, (_, amount) in zip(["X", "Y", "Z"], yearly_amounts):
            result[column] = amount

    return result


def find_financing_table(soup):
    for table in soup.find_all("table"):
        text = normalize_text(table.get_text(" ")).lower()
        has_year_column = re.search(r"на\s+\d{4}\s+год", text)
        if has_year_column and "всего" in text:
            return table

    return None


def parse_kbk(soup):
    result = {}
    table = find_kbk_table(soup)
    if not table:
        return result

    rows = table.find_all("tr")[1:]
    for row in rows:
        cells = [normalize_text(cell.get_text(" ")) for cell in row.find_all(["th", "td"])]
        if not cells or "Итого" in cells[0]:
            continue

        kbk = re.sub(r"\D", "", cells[0])
        if not kbk:
            continue

        target_column = classify_kbk(kbk)
        if target_column:
            result[target_column] = kbk

    return result


def find_kbk_table(soup):
    for table in soup.find_all("table"):
        text = normalize_text(table.get_text(" ")).lower()
        if "кбк" in text and re.search(r"на\s+20\d{2}\s+год", text):
            return table

    return find_table_by_headers(soup, ["КБК"])


def classify_kbk(kbk):
    if kbk in CHILD_KBK_CODES:
        return "AI"
    if kbk in ADULT_KBK_CODES:
        return "AH"

    if "1305920" in kbk:
        return "AI"

    return "AH"


def parse_medicine_table(soup):
    result = {}
    table = find_table_by_headers(soup, ["Мнн", "Количество", "Цена за ед."])
    if not table:
        return result

    rows = table.find_all("tr")
    if len(rows) < 2:
        return result

    cells = [normalize_text(cell.get_text(" ")) for cell in rows[1].find_all(["th", "td"])]
    cells = [cell for cell in cells if cell]
    if len(cells) >= 5:
        form_text = cells[1]
        quantity = parse_number(cells[2])
        unit_price = parse_number(cells[3])
        if quantity is not None:
            result["_eis_quantity"] = quantity
        if unit_price is not None:
            result["_eis_unit_price"] = unit_price
        result["_eis_form_text"] = form_text

    variant_table = find_table_by_headers(soup, ["Код позиции", "Цена за ед."])
    if variant_table:
        text = normalize_text(variant_table.get_text(" "))
        okpd_match = re.search(r"\b\d{2}\.\d{2}\.\d{2}\.\d{3}\b", text)
        if okpd_match:
            result["AJ"] = okpd_match.group(0)

    return result


def find_table_by_headers(soup, required_headers):
    for table in soup.find_all("table"):
        text = normalize_text(table.get_text(" ")).lower()
        if all(header.lower() in text for header in required_headers):
            return table
    return None
