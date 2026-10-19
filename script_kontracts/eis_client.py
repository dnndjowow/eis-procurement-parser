import time

import requests

from eis import parse_eis_by_notice_number
from settings import EIS_RETRY_ATTEMPTS, EIS_RETRY_DELAY_SECONDS


def has_eis_error(eis_data):
    values = eis_data.get("eis", {}).values()
    return any(str(value).startswith("Ошибка") for value in values)


def fetch_eis_data(notice_number, attempts=EIS_RETRY_ATTEMPTS):
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            eis_data = parse_eis_by_notice_number(notice_number)
            if not has_eis_error(eis_data):
                return eis_data

            last_error = RuntimeError('данные ЕИС не получены')
        except Exception as e:
            last_error = e

        if attempt < attempts:
            print(
                f'WARN {notice_number}: ошибка ЕИС, повтор {attempt}/{attempts} '
                f'({format_eis_error(last_error)})'
            )
            time.sleep(EIS_RETRY_DELAY_SECONDS)

    raise last_error


def format_eis_error(error):
    if isinstance(error, requests.exceptions.SSLError):
        return 'ошибка SSL-сертификата ЕИС'

    if isinstance(error, requests.exceptions.Timeout):
        return 'таймаут подключения к ЕИС'

    if isinstance(error, requests.exceptions.ConnectionError):
        return 'сайт ЕИС недоступен или нет интернета'

    if isinstance(error, requests.exceptions.HTTPError):
        response = error.response
        if response is not None:
            if response.status_code == 404:
                return 'контракт не найден в ЕИС'
            return f'HTTP {response.status_code} от ЕИС'
        return 'HTTP-ошибка ЕИС'

    message = str(error)
    if 'Не удалось найти контракт по номеру извещения' in message:
        return 'контракт не найден по номеру извещения'
    if 'Не удалось получить ссылку контракта' in message:
        return 'не удалось получить ссылку контракта'

    return message
