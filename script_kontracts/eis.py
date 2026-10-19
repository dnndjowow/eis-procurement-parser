import requests
from bs4 import BeautifulSoup
import json
import re
import sys
import os
from urllib.parse import urljoin

from settings import EIS_CA_BUNDLE

DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Referer': 'https://zakupki.gov.ru/',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Cache-Control': 'max-age=0',
}


def get_eis_verify():
    if EIS_CA_BUNDLE.exists():
        return str(EIS_CA_BUNDLE)

    return True


# Функция для извлечения названия поставщика
def extract_supplier_name(text):
    """Извлекает название поставщика только из скобок, убирая кавычки"""
    if not text:
        return "Не найдено"
    
    matches = re.findall(r'\((.*?)\)', text)
    
    if matches:
        supplier_text = matches[-1]
        supplier_text = supplier_text.replace('"', '').replace("'", "").strip()
        
        if '/' in supplier_text:
            supplier_text = supplier_text.split('/')[0].strip()
        
        return supplier_text
    return "Не найдено"

# УЛУЧШЕННАЯ Функция для очистки цены
def clean_price(price_text):
    """Убирает пробелы и лишний текст из цены контракта"""
    if price_text == "Не найдено":
        return price_text
    
    # 1. Убираем текст "Загрузка..." и другие нечисловые фрагменты
    price_text = re.sub(r'Загрузка\.\.\.', '', price_text, flags=re.IGNORECASE)
    price_text = re.sub(r'\s+', ' ', price_text)  # Заменяем множественные пробелы на один
    
    # 2. Ищем числовую часть цены (с запятой в качестве десятичного разделителя)
    price_match = re.search(r'[\d\s]+\d[\d\s]*,\d+', price_text)
    
    if price_match:
        # Берем только числовую часть
        price_text = price_match.group(0)
    else:
        # Альтернативный поиск: любой формат с числами
        price_match = re.search(r'[\d\s,\.]+', price_text)
        if price_match:
            price_text = price_match.group(0)
    
    # 3. Убираем все пробелы и лишние символы
    cleaned_price = price_text.replace(' ', '').replace('\xa0', '').replace('\u202f', '').replace('\u2009', '')
    
    # 4. Убираем лишние символы в начале и конце
    cleaned_price = cleaned_price.strip()
    
    return cleaned_price

def find_contract_urls_by_notice_number(notice_number, headers=None):
    """
    Ищет контракты в ЕИС по номеру извещения/закупки и возвращает ссылки на карточки.
    Логика поиска перенесена из parser_eis.py.
    """
    if headers is None:
        headers = DEFAULT_HEADERS

    search_url = (
        'https://zakupki.gov.ru/epz/contract/search/results.html?searchString='
        + str(notice_number)
        + '&morphology=on&fz44=on&contractStageList_0=on&contractStageList_1=on'
        + '&contractStageList_2=on&contractStageList_3=on&contractStageList=0%2C1%2C2%2C3'
        + '&budgetLevelsIdNameHidden=%7B%7D&sortBy=UPDATE_DATE&pageNumber=1'
        + '&sortDirection=false&recordsPerPage=_10&showLotsInfoHidden=false'
    )

    response = requests.get(search_url, headers=headers, timeout=30, verify=get_eis_verify())
    response.raise_for_status()
    response.encoding = 'utf-8'

    soup = BeautifulSoup(response.text, 'html.parser')
    contract_blocks = soup.find_all('div', class_='search-registry-entry-block box-shadow-search-input')

    if not contract_blocks:
        raise ValueError('Не удалось найти контракт по номеру извещения')

    contract_urls = []
    for block in contract_blocks:
        link = block.find('a', href=True)
        if link:
            href = link['href']
            if href.startswith('http'):
                contract_urls.append(href)
            else:
                contract_urls.append('https://zakupki.gov.ru' + href)

    if not contract_urls:
        raise ValueError('Не удалось получить ссылку контракта из результатов поиска')

    return contract_urls

def parse_eis_by_notice_number(notice_number, path_json=None, headers=None):
    """
    Ищет карточку контракта по номеру извещения/закупки и парсит первую найденную карточку.
    """
    contract_urls = find_contract_urls_by_notice_number(notice_number, headers=headers)
    return parse_eis_page(contract_urls[0], path_json=path_json, headers=headers)

# Основная функция для парсинга страницы
def parse_eis_page(url, path_json=None, headers=None):
    """
    Парсит страницу контракта EIS и возвращает данные в виде словаря
    
    Args:
        url (str): URL страницы контракта
        path_json (str, optional): Путь для сохранения JSON. Если None, не сохраняет.
        headers (dict, optional): Заголовки HTTP-запроса
    
    Returns:
        dict: Словарь с данными контракта
    """
    # Заголовки по умолчанию
    if headers is None:
        headers = DEFAULT_HEADERS
    
    try:
        # Отправляем GET-запрос к сайту
        response = requests.get(url, headers=headers, timeout=10, verify=get_eis_verify())
        response.raise_for_status()
        
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        page_url = response.url or url
        
        # 0. Номер реестровой записи контракта
        registry_number_value = "Не найдено"
        registry_link = soup.find('a', href=re.compile(r'/epz/contract/contractCard/common-info\.html\?reestrNumber='))
        
        if registry_link:
            page_url = urljoin('https://zakupki.gov.ru', registry_link.get('href', page_url))
            registry_text = registry_link.text.strip()
            # Убираем символ № и пробелы
            registry_number_value = registry_text.replace('№', '').replace(' ', '')
        else:
            # Альтернативный поиск через span с классом cardMainInfo_
            registry_span = soup.find('span', class_='cardMainInfo_')
            if registry_span:
                registry_link = registry_span.find('a')
                if registry_link:
                    registry_text = registry_link.text.strip()
                    registry_number_value = registry_text.replace('№', '').replace(' ', '')
                else:
                    registry_text = registry_span.text.strip()
                    registry_number_value = registry_text.replace('№', '').replace(' ', '')
        
        # 1. Дата заключения контракта
        date_conclusion = soup.find('span', class_='section__title', string='Дата заключения контракта')
        if date_conclusion:
            date_conclusion_value = date_conclusion.find_next('span', class_='section__info').text.strip()
        else:
            date_conclusion_value = "Не найдено"
            date_conclusion = soup.find(lambda tag: tag.name == 'span' and 'Дата заключения контракта' in tag.text)
            if date_conclusion:
                date_conclusion_value = date_conclusion.find_next('span', class_='section__info').text.strip()

        # 2. Цена контракта
        price = soup.find('span', class_='section__title', string='Цена контракта')
        if not price:
            price = soup.find(lambda tag: tag.name == 'span' and 'Цена контракта' in tag.text)
        
        price_value = "Не найдено"
        if price:
            price_info = price.find_next('span', class_='section__info')
            if price_info:
                # Получаем текст и очищаем
                raw_price = price_info.text.strip()
                price_value = clean_price(raw_price)
                
        # 3. Дата окончания исполнения контракта
        date_end = soup.find('span', class_='section__title', string='Дата окончания исполнения контракта')
        if not date_end:
            date_end = soup.find(lambda tag: tag.name == 'span' and 'Дата окончания исполнения контракта' in tag.text)
        
        date_end_value = "Не найдено"
        if date_end:
            date_end_info = date_end.find_next('span', class_='section__info')
            if date_end_info:
                date_end_value = date_end_info.text.strip()

        # 4. Номер контракта
        contract_number = soup.find('span', class_='section__title', string='Номер контракта')
        if not contract_number:
            contract_number = soup.find(lambda tag: tag.name == 'span' and 'Номер контракта' in tag.text)
        
        contract_number_value = "Не найдено"
        if contract_number:
            contract_info = contract_number.find_next('span', class_='section__info')
            if contract_info:
                contract_number_value = contract_info.text.strip()

        # 5. Информация о поставщиках
        supplier_name = "Не найдено"
        inn_value = "Не найдено"
        
        supplier_header = soup.find('h2', class_='blockInfo__title', string='Информация о поставщиках')
        
        if not supplier_header:
            supplier_header = soup.find(lambda tag: tag.name == 'h2' and 'поставщиках' in tag.text)
        
        if supplier_header:
            table = supplier_header.find_next('table', class_='blockInfo__table')
            
            if not table:
                table = supplier_header.find_next('table')
            
            if table:
                first_row = None
                tbody = table.find('tbody', class_='tableBlock__body')
                if tbody:
                    first_row = tbody.find('tr', class_='tableBlock__row')
                else:
                    first_row = table.find('tr', class_='tableBlock__row')
                
                if first_row:
                    first_cell = first_row.find('td', class_='tableBlock__col_first')
                    if not first_cell:
                        first_cell = first_row.find('td')
                    
                    if first_cell:
                        all_text = first_cell.get_text(" ", strip=True)
                        
                        supplier_name = extract_supplier_name(all_text)
                        
                        if supplier_name == "Не найдено":
                            parts = re.split(r'ИНН:|КПП:', all_text)
                            if parts:
                                supplier_name = parts[0].strip()
                        
                        inn_match = re.search(r'ИНН:\s*(\d+)', all_text)
                        if inn_match:
                            inn_value = inn_match.group(1)

        # Создаем структуру данных
        eis_data = {
            "eis": {
                "URL страницы": page_url,
                "Номер реестровой записи контракта": registry_number_value,
                "Дата заключения контракта": date_conclusion_value,
                "Цена контракта": price_value,
                "Дата окончания исполнения контракта": date_end_value,
                "Номер контракта": contract_number_value,
                "Поставщик": supplier_name,
                "ИНН": inn_value
            }
        }
        
        # Сохраняем в файл, если указан путь
        if path_json:
            # Создаем директорию если она не существует
            os.makedirs(os.path.dirname(path_json), exist_ok=True)
            
            with open(path_json, 'w', encoding='utf-8') as f:
                json.dump(eis_data, f, ensure_ascii=False, indent=2)
        
        return eis_data

    except requests.exceptions.RequestException as e:
        error_data = {
            "eis": {
                "URL страницы": url,
                "Номер реестровой записи контракта": "Ошибка подключения",
                "Дата заключения контракта": "Ошибка подключения",
                "Цена контракта": "Ошибка подключения",
                "Дата окончания исполнения контракта": "Ошибка подключения",
                "Номер контракта": "Ошибка подключения",
                "Поставщик": "Ошибка подключения",
                "ИНН": "Ошибка подключения"
            }
        }
        
        if path_json:
            with open(path_json, 'w', encoding='utf-8') as f:
                json.dump(error_data, f, ensure_ascii=False, indent=2)
        
        print(f"Ошибка при подключении к сайту: {e}")
        return error_data
        
    except Exception as e:
        error_data = {
            "eis": {
                "URL страницы": url,
                "Номер реестровой записи контракта": "Ошибка обработки",
                "Дата заключения контракта": "Ошибка обработки",
                "Цена контракта": "Ошибка обработки",
                "Дата окончания исполнения контракта": "Ошибка обработки",
                "Номер контракта": "Ошибка обработки",
                "Поставщик": "Ошибка обработки",
                "ИНН": "Ошибка обработки"
            }
        }
        
        if path_json:
            with open(path_json, 'w', encoding='utf-8') as f:
                json.dump(error_data, f, ensure_ascii=False, indent=2)
        
        print(f"Ошибка при обработке данных: {e}")
        import traceback
        traceback.print_exc()
        return error_data

# Если файл запускается напрямую (для тестирования)
if __name__ == "__main__":
    notice_number = "0873400003926000216"
    data = parse_eis_by_notice_number(notice_number)
    print(json.dumps(data, ensure_ascii=False, indent=2))
