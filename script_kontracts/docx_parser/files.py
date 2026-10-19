import json
import os
from typing import Dict, Tuple

from settings import FILES_DIR

from .contract import parse_contract
from .protocol import parse_protocol

DEFAULT_FILES_DIR = str(FILES_DIR)
DEFAULT_JSON_DIR = None


def find_files(files_dir: str = DEFAULT_FILES_DIR) -> Tuple[str, str]:
    """
    Находит файлы протокола и контракта в указанной директории.

    Returns:
        Кортеж: (путь_к_протоколу, путь_к_контракту)
    """
    if not os.path.exists(files_dir):
        raise FileNotFoundError(f"Директория не найдена: {files_dir}")

    files = [f for f in os.listdir(files_dir) if f.endswith('.docx')]

    if len(files) != 2:
        raise ValueError(f"В папке должно быть ровно 2 .docx файла, найдено: {len(files)}")

    def count_digits(filename: str) -> int:
        name_without_ext = os.path.splitext(filename)[0]
        return sum(c.isdigit() for c in name_without_ext)

    protocol_file = None
    contract_file = None

    for file in files:
        if 'ППИОП' in file or 'протокол' in file.lower():
            protocol_file = file
        elif 'Проект' in file or '-0001' in file:
            contract_file = file

    if not protocol_file or not contract_file:
        files_sorted = sorted(files, key=count_digits)
        protocol_file = files_sorted[0]
        contract_file = files_sorted[1]

    return (
        os.path.join(files_dir, protocol_file),
        os.path.join(files_dir, contract_file),
    )


def parse_document_pair(protocol_file: str, contract_file: str, json_dir: str = None) -> Tuple[Dict, Dict]:
    """
    Парсит конкретную пару файлов: протокол и контракт.

    Returns:
        Кортеж: (данные_протокола, данные_контракта)
    """
    protocol_data = parse_protocol(protocol_file)
    contract_data = parse_contract(contract_file)

    if json_dir:
        os.makedirs(json_dir, exist_ok=True)

        protocol_json_path = os.path.join(json_dir, "result.json")
        with open(protocol_json_path, 'w', encoding='utf-8') as f:
            json.dump(protocol_data, f, ensure_ascii=False, indent=2, default=str)

        contract_json_path = os.path.join(json_dir, "contract_data.json")
        with open(contract_json_path, 'w', encoding='utf-8') as f:
            json.dump(contract_data, f, ensure_ascii=False, indent=2, default=str)

    return protocol_data, contract_data


def parse_document_files(
    files_dir: str = DEFAULT_FILES_DIR,
    json_dir: str = DEFAULT_JSON_DIR,
) -> Tuple[Dict, Dict]:
    protocol_file, contract_file = find_files(files_dir)
    return parse_document_pair(protocol_file, contract_file, json_dir)
