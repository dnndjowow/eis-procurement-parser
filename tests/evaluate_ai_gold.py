"""Offline value comparison; no model calls and no spreadsheet writes."""

import argparse
from collections import Counter
from datetime import datetime
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import re
import unicodedata


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_GOLD = Path(__file__).resolve().parent / 'gold' / 'contracts_v1.json'
ABSENT = object()


def read_json(text):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f'Duplicate JSON key: {key}')
            result[key] = value
        return result

    def invalid(value):
        raise ValueError(f'Invalid JSON constant: {value}')

    return json.loads(text, object_pairs_hook=unique, parse_constant=invalid)


def locate(data, path):
    for key in path:
        if isinstance(key, int) and isinstance(data, list) and 0 <= key < len(data):
            data = data[key]
        elif isinstance(key, str) and isinstance(data, dict) and key in data:
            data = data[key]
        else:
            return ABSENT
    return data


def display_path(path):
    return '$' + ''.join(f'[{key}]' if isinstance(key, int) else f'.{key}' for key in path)


def normalized(value, kind):
    if value is None:
        return None
    if isinstance(value, bool):
        return ('invalid_boolean', value)
    if kind == 'number' and isinstance(value, (str, int, float)):
        text = str(value).strip()
        if re.fullmatch(r'\d+(?:[.,]\d+)?|\d{1,3}(?:[ \u00a0\u202f]\d{3})+(?:[.,]\d+)?', text):
            try:
                return Decimal(re.sub(r'\s', '', text).replace(',', '.'))
            except InvalidOperation:
                pass
        return ('invalid_number', text)
    if not isinstance(value, str):
        return ('invalid_type', repr(value))
    text = ' '.join(unicodedata.normalize('NFC', value).replace('ё', 'е').replace('Ё', 'Е').split())
    if kind == 'unit':
        aliases = {'руб': 'руб.', 'руб.': 'руб.', 'rub': 'руб.', 'rur': 'руб.',
                   'шт': 'шт.', 'шт.': 'шт.', 'штука': 'шт.', 'штука*': 'шт.', 'штук': 'шт.',
                   'упак': 'упак.', 'упак.': 'упак.', 'упк.': 'упак.', 'уп.': 'упак.',
                   'упаковки': 'упак.', 'упаковка': 'упак.', 'упаковок': 'упак.',
                   'ме': 'МЕ', 'мг': 'мг', 'мл': 'мл'}
        return '/'.join(aliases.get(p.strip().lower(), p.strip()) for p in text.split('/'))
    return text


def compare_part(actual, expected, fixed_currency=False):
    rows = []
    for fact in expected['facts']:
        if fixed_currency and fact['path'] == ['currency']:
            continue
        value = locate(actual, fact['path'])
        target = fact['value']
        if value is ABSENT:
            status = 'MISSING'
        elif target is None and value is None:
            status = 'EXPECTED_NULL'
        elif value is None:
            status = 'MISSING'
        elif normalized(value, fact['kind']) == normalized(target, fact['kind']):
            status = 'CORRECT'
        else:
            status = 'WRONG'
        rows.append({'path': display_path(fact['path']), 'status': status, 'expected': target,
                     'actual': None if value is ABSENT else value, 'field_absent': value is ABSENT,
                     'sources': fact['sources'], 'note': fact.get('note', ''), 'check': 'value'})
    for shape in expected['lists']:
        value = locate(actual, shape['path'])
        if not isinstance(value, list):
            status = 'MISSING' if value is ABSENT or value is None else 'WRONG'
            count = None
        else:
            count = len(value)
            status = 'CORRECT' if count == shape['length'] else 'EXTRA' if count > shape['length'] else 'MISSING'
        rows.append({'path': display_path(shape['path']) + '.length', 'status': status,
                     'expected': shape['length'], 'actual': count, 'sources': [], 'note': '', 'check': 'list'})
        if isinstance(value, list):
            for index in range(shape['length'], len(value)):
                rows.append({'path': display_path(shape['path'] + [index]), 'status': 'EXTRA',
                             'expected': None, 'actual': value[index], 'sources': [],
                             'note': 'Лишняя запись отсутствует в эталоне.', 'check': 'extra_record'})
    return rows


def three_block_reference(reference):
    # Manually verified technical variants in the two reference Word documents.
    positions = {
        '0873400003926000653-0001 Проект.docx': [0, 1, 2],
        '0873400003926000666-0001.docx': [0],
    }
    selected = positions.get(reference['file'])
    if selected is None:
        raise ValueError('Для трёх блоков нужен отдельный проверенный эталон технических вариантов')
    original = reference['parts']['1']
    specification = {
        'facts': [f for f in original['facts'] if 'technical_variants' not in f['path'] and f['path'] != ['currency']],
        'lists': [s for s in original['lists'] if 'technical_variants' not in s['path']],
    }
    technical = {'facts': [], 'lists': [{'path': ['technical_variants'], 'length': len(selected)}]}
    for target, position in enumerate(selected):
        prefix = ['products', position, 'technical_variants', 0]
        for kind in ('facts', 'lists'):
            for record in original[kind]:
                if record['path'][:len(prefix)] == prefix:
                    technical[kind].append({**record, 'path': ['technical_variants', target] + record['path'][len(prefix):]})
    return {'1': specification, '2': technical, '3': reference['parts']['2']}


def evaluate_run(run_dir, folder=BASE_DIR / 'test_ai_contracts', gold_path=DEFAULT_GOLD):
    run_dir, folder = Path(run_dir), Path(folder)
    settings = read_json((run_dir / 'settings.json').read_text(encoding='utf-8'))
    is_three = settings.get('schema_version') == 'three_blocks_v1'
    if not is_three and not settings.get('schema_version', '').startswith('split_compact'):
        raise ValueError('Нужен прогон --benchmark-split с компактной схемой. Старый полный JSON не смешивается с новым.')
    gold = read_json(Path(gold_path).read_text(encoding='utf-8'))
    references = {doc['file']: doc for doc in gold['documents']}
    output = run_dir / datetime.now().strftime('gold_check_%Y-%m-%d_%H-%M-%S_%f')
    output.mkdir()
    lines = ['СРАВНЕНИЕ С ЭТАЛОНОМ WORD',
             'Оценка только перечисленных в эталоне полей; это не точность на всех контрактах.',
             'Числа сравниваются без округления; пробелы, е/ё и известные единицы нормализуются.',
             'Позиции и варианты сопоставляются по порядку документа. Повторы сохраняются.',
             'Формулировки сроков, технические ID и ссылки моделей здесь не оцениваются.',
             'Сохранённый статус валидатора показан отдельно: правильное значение не подтверждает его источник.',
             'EXPECTED_NULL учитывается отдельно и не увеличивает число правильно извлечённых значений.',
             'Отказы API и неразобранный JSON не исключаются из количества попыток.']
    results = []
    fixed_currency = is_three or settings.get('schema_version') == 'split_compact_v3'
    if fixed_currency:
        lines.append('Валюта контракта задана кодом (руб.): не оценивается как извлечение модели.')
    for index, name in enumerate(settings['documents'], 1):
        lines.append(f'\nДОКУМЕНТ: {name}')
        reference = references.get(name)
        path = folder / name
        if path.name != name or reference is None or not path.is_file():
            lines.append('BLOCKED: нет эталона или исходного Word; оценка не выполняется.')
            results.append({'file': name, 'evaluation': 'BLOCKED', 'reason': 'missing_document_or_gold'})
            continue
        if hashlib.sha256(path.read_bytes()).hexdigest() != reference['sha256']:
            lines.append('BLOCKED: SHA-256 Word изменился; эталон нужно пересмотреть.')
            results.append({'file': name, 'evaluation': 'BLOCKED', 'reason': 'document_hash_mismatch'})
            continue
        if is_three:
            try:
                reference = {**reference, 'parts': three_block_reference(reference)}
            except ValueError as error:
                lines.append('BLOCKED: ' + str(error))
                results.append({'file': name, 'evaluation': 'BLOCKED', 'reason': 'missing_three_block_gold'})
                continue
        attempt_paths = sorted((run_dir / f'{index:03d}').glob('attempt_*.json'))
        if not attempt_paths:
            lines.append('NO_ATTEMPTS: сохранённых ответов нет.')
        for attempt_path in attempt_paths:
            attempt = read_json(attempt_path.read_text(encoding='utf-8'))
            part = str(attempt['part'])
            result = {'file': name, 'model': attempt['model'], 'part': int(part),
                      'seconds': attempt.get('seconds'), 'saved_status': attempt['status'],
                      'errors': attempt.get('errors', []), 'checks': []}
            lines.append(f"\nМОДЕЛЬ: {attempt['model']}; блок {part}; сохранённый статус {attempt['status']}; {attempt.get('seconds', 0):.2f} с")
            raw = attempt.get('raw_response')
            if raw is None:
                result['evaluation'] = 'NO_RESPONSE'
                lines.append('NO_RESPONSE: данные не получены; оценить правильность значений нельзя.')
            else:
                try:
                    wrapper = re.fullmatch(r'\s*```(?:json)?\s*\n([\s\S]*?)\n```\s*', raw)
                    actual = read_json(wrapper.group(1) if wrapper else raw)
                    if not isinstance(actual, dict):
                        raise ValueError('Ответ не является JSON-объектом')
                    result['checks'] = compare_part(actual, reference['parts'][part], fixed_currency=fixed_currency)
                    bad = any(row['status'] in ('MISSING', 'WRONG', 'EXTRA') for row in result['checks'])
                    result['evaluation'] = 'DIFFERENCES' if bad else 'VALUES_MATCH'
                    for row in result['checks']:
                        lines.append(f"{row['status']} {row['path']}: эталон={row['expected']!r}; ответ="
                                     + ('поле отсутствует' if row.get('field_absent') else repr(row['actual'])))
                        if row['sources']:
                            lines.append('  Источник эталона: ' + ', '.join(row['sources']))
                        if row['note']:
                            lines.append('  ' + row['note'])
                except (ValueError, TypeError) as error:
                    result['evaluation'] = 'INVALID_JSON'
                    result['parse_error'] = str(error)
                    lines.append(f'INVALID_JSON: {error}')
            results.append(result)
    lines.append('\nСВОДКА ПО МОДЕЛЯМ И БЛОКАМ')
    groups = {}
    for result in results:
        if 'model' in result:
            groups.setdefault((result['model'], result['part']), []).append(result)
    for (model, part), attempts in groups.items():
        counts = Counter(a['evaluation'] for a in attempts)
        fields = Counter(row['status'] for a in attempts for row in a['checks'] if row['check'] == 'value')
        shapes = Counter(row['status'] for a in attempts for row in a['checks'] if row['check'] == 'list')
        lines.append(f'\n{model}, блок {part}: попыток {len(attempts)}; результаты {dict(counts)}')
        lines.append(f'  Значения: {dict(fields)}; количество записей: {dict(shapes)}')
    (output / 'results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    (output / 'summary.txt').write_text('\n'.join(lines) + '\n', encoding='utf-8')
    print(f'Эталонная проверка сохранена: {output / "summary.txt"}')
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Сравнить сохранённые компактные ответы с эталоном, без API')
    parser.add_argument('run', type=Path, help='Папка reports/ai_split_...')
    parser.add_argument('--folder', type=Path, default=BASE_DIR / 'test_ai_contracts')
    args = parser.parse_args()
    evaluate_run(args.run, args.folder)
