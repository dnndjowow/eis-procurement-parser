import re
from functools import lru_cache
from pathlib import Path

from docx import Document

from docx_parser.common import collect_document_lines, normalize_text
from settings import FILES_DIR


def is_protocol_file(path):
    filename = path.name.lower()
    return 'ппиоп' in filename or 'протокол' in filename


def is_docx_file(path):
    return (
        path.suffix.lower() == '.docx'
        and not path.name.startswith('~$')
        and not path.name.startswith('._')
    )


@lru_cache(maxsize=512)
def read_docx_text(path_string):
    try:
        doc = Document(path_string)
    except Exception:
        return ""

    return "\n".join(collect_document_lines(doc))


def document_contains_notice(path, notice_number):
    text = read_docx_text(str(path))
    return notice_number in re.sub(r'\D', '', text)


def is_protocol_document(path):
    if is_protocol_file(path):
        return True

    text = normalize_text(read_docx_text(str(path))).lower()[:3000]
    return (
        'протокол' in text
        and (
            'подведения итогов' in text
            or 'рассмотрения заявок' in text
            or 'запроса котировок' in text
        )
    )


def is_contract_document(path):
    filename = path.name.lower()
    if 'ппиоп' in filename or 'протокол' in filename:
        return False

    text = normalize_text(read_docx_text(str(path))).lower()[:3000]
    return (
        'государственный контракт' in text
        or 'контракт №' in text
        or 'контракт N' in text
        or 'проект государственного контракта' in text
    )


def append_unique(paths, path):
    if path not in paths:
        paths.append(path)


def find_documents_by_content(notice_number, files_dir, known_files):
    content_files = []

    for path in files_dir.glob('*.docx'):
        if not is_docx_file(path) or path in known_files:
            continue
        if document_contains_notice(path, notice_number):
            content_files.append(path)

    return content_files


def find_document_files(notice_number, files_dir=FILES_DIR):
    files_dir = Path(files_dir)
    matched_files = [
        path
        for path in files_dir.glob(f'*{notice_number}*.docx')
        if is_docx_file(path)
    ]

    protocol_files = [path for path in matched_files if is_protocol_file(path)]
    contract_files = [path for path in matched_files if not is_protocol_file(path)]

    if len(protocol_files) != 1 or len(contract_files) != 1:
        content_files = find_documents_by_content(notice_number, files_dir, set(matched_files))
        for path in content_files:
            append_unique(matched_files, path)
            if is_contract_document(path):
                append_unique(contract_files, path)
            elif is_protocol_document(path):
                append_unique(protocol_files, path)

    return matched_files, protocol_files, contract_files


def describe_document_problem(matched_files, protocol_files, contract_files):
    if not matched_files:
        return 'найдено 0/2, нет контракта и протокола'

    details = []
    if len(contract_files) == 0:
        details.append('нет контракта')
    elif len(contract_files) > 1:
        details.append(f'контрактов {len(contract_files)}')

    if len(protocol_files) == 0:
        details.append('нет протокола')
    elif len(protocol_files) > 1:
        details.append(f'протоколов {len(protocol_files)}')

    found_count = min(len(contract_files), 1) + min(len(protocol_files), 1)
    message = f'найдено {found_count}/2'

    if details:
        message += ', ' + ', '.join(details)

    if len(matched_files) > 2:
        message += f', всего файлов с номером: {len(matched_files)}'

    return message


def find_document_pair(notice_number, files_dir=FILES_DIR):
    matched_files, protocol_files, contract_files = find_document_files(notice_number, files_dir)

    if not matched_files:
        return None

    if len(protocol_files) != 1 or len(contract_files) != 1:
        problem = describe_document_problem(matched_files, protocol_files, contract_files)
        file_names = ', '.join(path.name for path in matched_files)
        raise ValueError(f'{problem}; файлы: {file_names}')

    return protocol_files[0], contract_files[0]
