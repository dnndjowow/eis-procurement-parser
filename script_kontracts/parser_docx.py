import json
from docx_parser import (
    ContractParser,
    DEFAULT_FILES_DIR,
    DEFAULT_JSON_DIR,
    find_files,
    parse_contract,
    parse_delivery_terms,
    parse_document_files,
    parse_document_pair,
    parse_protocol,
)

__all__ = [
    "ContractParser",
    "DEFAULT_FILES_DIR",
    "DEFAULT_JSON_DIR",
    "find_files",
    "parse_contract",
    "parse_delivery_terms",
    "parse_document_files",
    "parse_document_pair",
    "parse_protocol",
]

# ================== ТОЧКА ВХОДА ==================
def main():
    """Основная функция для запуска из командной строки"""
    try:
        # Используем конфигурацию по умолчанию
        protocol_data, contract_data = parse_document_files()
        
        print("\nРезультаты парсинга:")
        print("\nДанные протокола:")
        print(json.dumps(protocol_data, ensure_ascii=False, indent=2, default=str))
        print("\nДанные контракта:")
        print(json.dumps(contract_data, ensure_ascii=False, indent=2, default=str))
        
    except FileNotFoundError as e:
        print(f"Файл не найден: {e}")
    except ValueError as e:
        print(f"Ошибка данных: {e}")
    except Exception as e:
        print(f"Неизвестная ошибка: {e}")

if __name__ == "__main__":
    main()
