import contextlib
import copy
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from docx import Document


SPEC = importlib.util.spec_from_file_location('ai_experiment', Path(__file__).resolve().parents[1] / 'test.py')
ai = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(ai)


def fixture(schema):
    if isinstance(schema, list):
        return [] if schema[0] == 'warning' else [fixture(schema[0])]
    if isinstance(schema, dict):
        result = {key: fixture(value) for key, value in schema.items()}
        if 'packages_unit' in result:
            result['packages_unit'] = 'упак.'
        if 'price' in result:
            result['unit'] = 'руб./шт.'
        fields = [key for key, value in schema.items() if isinstance(value, str)]
        if fields:
            result['sources'] = {key: ['DOC/P1'] for key in fields}
        return result
    return '2027-01-11' if schema == 'date' else '10' if 'number' in schema or schema == 'integer' else 'text'


def record_fixture(node):
    if isinstance(node, list):
        return [record_fixture(item) for item in node]
    if not isinstance(node, dict):
        return node
    result = {k: record_fixture(v) for k, v in node.items()
              if k not in ('sources', 'currency', 'delivery_text', 'document_text')}
    if 'sources' in node:
        result['sources'] = list(dict.fromkeys(r for refs in node['sources'].values() for r in refs))
    return result


class CompactTests(unittest.TestCase):
    def validate(self, data, part=1):
        notices, corrections = [], []
        with patch.object(ai, 'check_extraction_evidence'):
            parsed, errors = ai.validate_simple_answer(json.dumps(data), part, {'DOC/P1': 'text'}, {}, corrections, notices)
        return parsed, errors, notices

    def test_valid_both_schemas(self):
        for part in (0, 1):
            data = fixture(ai.SPLIT_SCHEMAS[part])
            if part == 1:
                data['unassigned_document_deadlines'] = []
            _, errors, notices = self.validate(data, part)
            self.assertEqual(errors, [])
            self.assertEqual(notices, [])

    def test_missing_category_is_not_an_error(self):
        data = fixture(ai.SPLIT_SCHEMAS[1])
        q = data['stages'][0]['quantities'][0]
        q['children'] = q['adults'] = None
        del q['sources']['children'], q['sources']['adults']
        _, errors, notices = self.validate(data)
        self.assertEqual(errors, [])
        self.assertFalse(any('children' in n or 'adults' in n for n in notices))

    def test_invalid_date_unit_and_source(self):
        data = fixture(ai.SPLIT_SCHEMAS[1])
        stage = data['stages'][0]
        stage['delivery_deadline'] = '2027-02-30'
        stage['sources']['stage'] = ['DOC/P999']
        stage['quantities'][0]['unit'] = None
        stage['quantities'][0]['sources'].pop('unit')
        _, errors, _ = self.validate(data)
        self.assertTrue(any('существующая дата' in e for e in errors))
        self.assertTrue(any('нужна единица' in e for e in errors))
        self.assertTrue(any('координаты отсутствуют' in e for e in errors))

    def test_missing_source_and_unknown_key(self):
        data = fixture(ai.SPLIT_SCHEMAS[0])
        data['sources'].pop('contract_price')
        data['product_id'] = 'unexpected'
        _, errors, _ = self.validate(data, 0)
        self.assertTrue(any('источник' in e for e in errors))
        self.assertTrue(any('неизвестное поле' in e for e in errors))

    def test_repeated_variants_remain(self):
        data = fixture(ai.SPLIT_SCHEMAS[0])
        variants = data['products'][0]['technical_variants']
        variants.append(copy.deepcopy(variants[0]))
        parsed, errors, _ = self.validate(data, 0)
        self.assertEqual(errors, [])
        self.assertEqual(len(parsed['products'][0]['technical_variants']), 2)

    def test_unassigned_deadline_is_not_filled_in(self):
        data = fixture(ai.SPLIT_SCHEMAS[1])
        stage = data['stages'][0]
        stage['document_deadline'] = None
        stage['sources'].pop('document_deadline')
        parsed, errors, notices = self.validate(data)
        self.assertEqual(errors, [])
        self.assertIsNone(parsed['stages'][0]['document_deadline'])
        self.assertTrue(any(n.startswith('REVIEW') for n in notices))

    def test_duplicate_json_key_rejected(self):
        _, errors = ai.validate_simple_answer('{"warnings":[],"warnings":[]}', 0, {}, {}, [], [])
        self.assertIn('Повторяющийся ключ', errors[0])

    def test_alternation_and_pair_reports_without_api(self):
        calls = []

        def ask(model, prompt):
            part = int(model) % 2
            self.assertTrue(prompt.endswith('BLOCK' + str(part)))
            calls.append(model)
            data = fixture(ai.SPLIT_SCHEMAS[part])
            if part == 1:
                data['unassigned_document_deadlines'] = []
            return json.dumps(record_fixture(data))

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            folder = root / 'docs'
            folder.mkdir()
            Document().save(folder / 'sample.docx')
            blocks = [('BLOCK0', {'DOC/P1': 'text'}, {}), ('BLOCK1', {'DOC/P1': 'text'}, {})]
            with patch.object(ai, 'TEST_MODELS', ['0', '1', '2', '3']), \
                 patch.object(ai, 'split_document_inputs', return_value=blocks), \
                 patch.object(ai, 'check_extraction_evidence'), \
                 patch.object(ai, 'select_record_sources', side_effect=lambda field, refs, *args: refs), \
                 patch.object(ai, 'ask_model_with_deadline', side_effect=ask), \
                 contextlib.redirect_stdout(io.StringIO()):
                ai.run_split_benchmark(folder, root / 'reports')
            output = next((root / 'reports').iterdir())
            for path in (output / '001').glob('pair_*.json'):
                pair = json.loads(path.read_text())
                self.assertTrue(pair['complete'])
                self.assertFalse(pair['requires_review'])
                self.assertIn('products', pair['data'])
                self.assertIn('stages', pair['data'])
                self.assertEqual(pair['data']['currency'], 'руб.')
            self.assertEqual(len(list((output / '001').glob('pair_*.json'))), 2)
            self.assertIn('Источник DOC/P1', (output / 'summary.txt').read_text())
        self.assertEqual(calls, ['0', '1', '2', '3'])

    def test_redundant_group_sources_are_audited(self):
        data = fixture(ai.SPLIT_SCHEMAS[1])
        data['stages'][0]['sources']['quantities'] = ['DOC/P1']
        corrections = []
        with patch.object(ai, 'check_extraction_evidence'):
            parsed, errors = ai.validate_simple_answer(json.dumps(data), 1, {'DOC/P1': 'text'}, {}, corrections, [])
        self.assertEqual(errors, [])
        self.assertNotIn('quantities', parsed['stages'][0]['sources'])
        self.assertTrue(any('дублирующая' in c['reason'] for c in corrections))

    def test_nonredundant_group_source_is_rejected(self):
        data = fixture(ai.SPLIT_SCHEMAS[1])
        data['stages'][0]['sources']['quantities'] = ['DOC/P2']
        with patch.object(ai, 'check_extraction_evidence'):
            _, errors = ai.validate_simple_answer(json.dumps(data), 1, {'DOC/P1': 'text', 'DOC/P2': 'other'}, {}, [], [])
        self.assertTrue(any('не подтверждена' in e for e in errors))


class EvidenceTests(unittest.TestCase):
    def setUp(self):
        self.sources = {
            'DOC/P1': '9.4. Документы не ранее «15» января 2027 г. и не позднее «15» февраля 2027 г.',
            'DOC/T1/R1C1': 'Этап поставки', 'DOC/T1/R1C2': 'Срок поставки',
            'DOC/T1/R1C3': 'Количество Товара, МЕ, Дети',
            'DOC/T1/R1C4': 'Количество Товара, МЕ, Взрослые',
            'DOC/T1/R2C1': 'I этап', 'DOC/T1/R2C2': 'не позднее 11.01.2027',
            'DOC/T1/R2C3': '48 000', 'DOC/T1/R2C4': '1 921 000',
        }
        self.data = {'stages': [{
            'stage': 'I этап', 'delivery_deadline': '2027-01-11', 'delivery_text': 'не позднее 11.01.2027',
            'document_deadline': '2027-02-15', 'document_text': self.sources['DOC/P1'],
            'quantities': [{'unit': 'МЕ', 'children': '48000', 'adults': '1921000', 'total': None,
                            'sources': {'unit': ['DOC/T1/R1C3'], 'children': ['DOC/T1/R2C3'], 'adults': ['DOC/T1/R2C4']}}],
            'sources': {'stage': ['DOC/T1/R2C1'], 'delivery_deadline': ['DOC/T1/R2C2'],
                        'delivery_text': ['DOC/T1/R2C2'], 'document_deadline': ['DOC/P1'], 'document_text': ['DOC/P1']},
        }], 'unassigned_document_deadlines': [], 'warnings': []}

    def validate(self):
        notices = []
        _, errors = ai.validate_simple_answer(json.dumps(self.data), 1, self.sources, {}, [], notices)
        return errors, notices

    def test_verified_stage_passes(self):
        self.assertEqual(self.validate(), ([], []))

    def test_lower_date_is_not_deadline(self):
        self.data['stages'][0]['document_deadline'] = '2027-01-15'
        self.assertTrue(any('конечный срок' in e for e in self.validate()[0]))

    def test_date_only_delivery_cell(self):
        self.sources['DOC/T1/R2C2'] = '11.01.2027'
        self.data['stages'][0]['delivery_text'] = '11.01.2027'
        self.assertEqual(self.validate(), ([], []))

    def test_document_date_of_another_stage(self):
        self.sources['DOC/P1'] = ('9.4. не позднее «15» февраля 2027 г. - по первому этапу поставки '
                                 'и не позднее «01» апреля 2027 г. - по второму этапу поставки')
        self.data['stages'][0]['document_deadline'] = '2027-04-01'
        self.assertTrue(any('другому этапу' in e for e in self.validate()[0]))

    def test_changed_quantity_is_rejected(self):
        self.data['stages'][0]['quantities'][0]['children'] = '48001'
        self.assertTrue(any('не найдено' in e for e in self.validate()[0]))

    def test_omitted_stage_is_rejected(self):
        self.data['stages'] = []
        self.assertTrue(any('пропущены этапы' in e for e in self.validate()[0]))

    def test_wrong_package_row_and_country_section(self):
        sources = {'DOC/T1/R1C1': 'Торговое наименование', 'DOC/T1/R1C2': 'Препарат',
                   'DOC/T1/R2C1': 'Форма выпуска', 'DOC/T1/R2C2': 'шприц 1, салфетки 2',
                   'DOC/T1/R3C1': 'Страна', 'DOC/T1/R3C2': 'Российская Федерация'}
        data = {'quantity': '1', 'unit': 'шприц', 'country': 'Российская Федерация',
                'sources': {'quantity': ['DOC/T1/R2C2'], 'country': ['DOC/T1/R3C2']}}
        errors = []
        ai.check_extraction_evidence(data, 0, sources, errors, [])
        self.assertEqual(len(errors), 1)

    def test_contract_price_mismatch(self):
        errors = []
        ai.check_extraction_evidence({'contract_price': '20', 'sources': {'contract_price': ['DOC/P1']}}, 0,
            {'DOC/P1': '2.2. Цена Контракта составляет 10 (Десять) руб. 00 коп.'}, errors, [])
        self.assertTrue(any('цена не совпадает' in e for e in errors))


class RecordSourcesTests(EvidenceTests):
    def test_explanations_do_not_trigger_review(self):
        data = record_fixture(self.data)
        data['explanations'] = ['Использована верхняя граница интервала.']
        notices = []
        _, errors = ai.validate_record_answer(json.dumps(data), 1, self.sources, {}, [], notices)
        self.assertEqual((errors, notices), ([], []))
        data['warnings'] = ['Неоднозначный срок.']
        ai.validate_record_answer(json.dumps(data), 1, self.sources, {}, [], notices)
        self.assertTrue(any(n.startswith('REVIEW') for n in notices))

    def validate(self):
        notices = []
        self.parsed, errors = ai.validate_record_answer(
            json.dumps(record_fixture(self.data)), 1, self.sources, {}, [], notices)
        return errors, notices

    def test_shared_sources_do_not_allow_swapped_categories(self):
        q = self.data['stages'][0]['quantities'][0]
        q['children'], q['adults'] = q['adults'], q['children']
        errors, _ = self.validate()
        self.assertTrue(any('.children' in e and 'не найдено' in e for e in errors))
        self.assertTrue(any('.adults' in e and 'не найдено' in e for e in errors))

    def test_nonexistent_record_source_is_rejected(self):
        self.data['stages'][0]['sources']['stage'].append('DOC/P999')
        self.assertTrue(any('координаты отсутствуют' in e for e in self.validate()[0]))

    def test_output_preserves_record_lists(self):
        self.assertEqual(self.validate(), ([], []))
        stage = self.parsed['stages'][0]
        self.assertIsInstance(stage['sources'], list)
        self.assertNotIn('delivery_text', stage)
        self.assertNotIn('document_text', stage)
        self.assertIsInstance(stage['quantities'][0]['sources'], list)

    def test_unknown_mapping_is_not_silently_accepted(self):
        data = record_fixture(self.data)
        data['stages'][0]['sources'] = {'stage': ['DOC/T1/R2C1']}
        _, errors = ai.validate_record_answer(json.dumps(data), 1, self.sources, {}, [], [])
        self.assertTrue(any('один список' in e for e in errors))


class RepairTests(unittest.TestCase):
    def test_source_patch_cannot_change_values(self):
        attempt = {'model': 'offline', 'part': 1, 'status': 'INVALID', 'seconds': 1,
                   'raw_response': '{}', 'parsed': {'contract_price': '10', 'sources': ['DOC/P9']},
                   'errors': ['$.sources[0]: координаты отсутствуют'], 'notices': []}
        patch_data = {'updates': [{'path': '$', 'sources': ['DOC/P1']}]}
        with patch.object(ai, 'ask_model_with_deadline', return_value=json.dumps(patch_data)), \
             patch.object(ai, 'validate_record_answer', return_value=({}, [])):
            result = ai.repair_sources_only('offline', 0, attempt, {'DOC/P1': '2.2. Цена 10 руб.'}, {})
        self.assertEqual(json.loads(result['raw_response'])['contract_price'], '10')
        self.assertEqual(attempt['parsed']['sources'], ['DOC/P9'])
        self.assertEqual(result['status'], 'OK')
        patch_data['updates'][0]['contract_price'] = '20'
        with patch.object(ai, 'ask_model_with_deadline', return_value=json.dumps(patch_data)):
            result = ai.repair_sources_only('offline', 0, attempt, {'DOC/P1': '2.2. Цена 10 руб.'}, {})
        self.assertEqual(result['status'], 'INVALID')
        self.assertEqual(result['parsed']['contract_price'], '10')

    def test_source_repair_not_used_for_value_errors(self):
        attempt = {'parsed': {'price': '10', 'sources': []},
                   'errors': ['$.price: число не совпадает'], 'status': 'INVALID'}
        with patch.object(ai, 'ask_model_with_deadline') as request:
            self.assertIs(ai.repair_sources_only('offline', 0, attempt, {}, {}), attempt)
        request.assert_not_called()

    def test_unit_errors_are_reported_with_bad_sources(self):
        data = record_fixture(fixture(ai.SPLIT_SCHEMAS[0]))
        product = data['products'][0]
        product['packages_unit'] = 'шт.'
        product['prices'][0]['unit'] = 'руб./единица товара'
        data['sources'] = ['DOC/P999']
        _, errors = ai.validate_record_answer(json.dumps(data), 0, {}, {}, [], [])
        self.assertTrue(any('координаты отсутствуют' in e for e in errors))
        self.assertTrue(any('Нужна единица упаковок' in e for e in errors))
        self.assertTrue(any('Нужна конкретная единица цены' in e for e in errors))

    def test_broad_sources_removed_only_with_valid_concrete_evidence(self):
        data = record_fixture(fixture(ai.SPLIT_SCHEMAS[0]))
        data['sources'].append('DOC/T1')
        sources = {'DOC/P1': 'text', 'DOC/T1/R1C1': 'table'}
        corrections = []
        with patch.object(ai, 'check_extraction_evidence'), \
             patch.object(ai, 'select_record_sources', side_effect=lambda key, refs, *args: refs):
            parsed, errors = ai.validate_record_answer(json.dumps(data), 0, sources, {}, corrections, [])
            self.assertEqual(errors, [])
            self.assertEqual(parsed['sources'], ['DOC/P1'])
            self.assertTrue(any('общие ссылки' in c['reason'] for c in corrections))
            data['sources'] = ['DOC/T1']
            _, errors = ai.validate_record_answer(json.dumps(data), 0, sources, {}, [], [])
            self.assertTrue(any('общая ссылка' in e for e in errors))

    def test_expected_country_null_requires_tech_without_country_field(self):
        data = record_fixture(fixture(ai.SPLIT_SCHEMAS[0]))
        country = data['products'][0]['technical_variants'][0]['countries_of_origin'][0]
        country.update(country=None, sources=[])
        sources = {'DOC/P1': 'text', 'DOC/T1/R1C1': 'Торговое наименование'}
        with patch.object(ai, 'check_extraction_evidence'), \
             patch.object(ai, 'select_record_sources', side_effect=lambda key, refs, *args: refs):
            notices = []
            ai.validate_record_answer(json.dumps(data), 0, sources, {}, [], notices)
            self.assertFalse(any('.country:' in n for n in notices))
            sources['DOC/T1/R1C2'] = 'Страна происхождения Товара'
            notices = []
            ai.validate_record_answer(json.dumps(data), 0, sources, {}, [], notices)
            self.assertTrue(any('.country:' in n for n in notices))

    def test_products_only_never_requests_stages(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            folder = root / 'docs'
            folder.mkdir()
            Document().save(folder / 'sample.docx')
            result = {'model': 'fin', 'part': 1, 'status': 'INVALID',
                      'errors': ['offline'], 'notices': [], 'seconds': 0}
            with patch.object(ai, 'TEST_MODELS', ['fin', 'sante']), \
                 patch.object(ai, 'split_document_inputs', return_value=[('ONE', {}, {}), ('TWO', {}, {})]), \
                 patch.object(ai, 'extract_split_attempt', return_value=result) as extract, \
                 contextlib.redirect_stdout(io.StringIO()):
                ai.run_split_benchmark(folder, root / 'reports', products_only=True)
            self.assertEqual(extract.call_count, 1)
            self.assertEqual(extract.call_args.args[0:2], ('fin', 0))
            self.assertEqual(extract.call_args.args[3], 'ONE')
            self.assertEqual(list((root / 'reports').glob('*/001/pair_*.json')), [])

    def test_price_unit_and_package_unit_are_checked(self):
        sources = {'DOC/T1/R1C1': 'Цена за единицу (мг), итого, руб.',
                   'DOC/T1/R2C1': '10', 'DOC/T1/R1C2': 'Количество вторичных упаковок по потребности'}
        data = {'price': '10', 'unit': 'руб./мл', 'packages_unit': 'шт',
                'sources': {'price': ['DOC/T1/R2C1'], 'unit': ['DOC/T1/R1C1']}}
        errors = []
        ai.check_extraction_evidence(data, 0, sources, errors, [])
        self.assertTrue(any('противоречит заголовку' in e for e in errors))
        self.assertTrue(any('в упаковках' in e for e in errors))
        data['unit'], data['packages_unit'] = 'руб./мг', 'упак.'
        errors, notices = [], []
        ai.check_extraction_evidence(data, 0, sources, errors, notices)
        self.assertEqual((errors, notices), ([], []))

    def test_generic_price_unit_is_rejected(self):
        errors = []
        ai.check_extraction_evidence({'price': '10', 'unit': 'руб./единица товара', 'sources': {}},
                                    0, {}, errors, [])
        self.assertTrue(any('не конкретная единица' in e for e in errors))

    def test_single_repair_includes_input_and_errors(self):
        with patch.object(ai, 'ask_model_with_deadline', side_effect=['broken', '{}']) as request, \
             patch.object(ai, 'validate_record_answer', side_effect=[(None, ['bad JSON']), ({}, [])]), \
             contextlib.redirect_stdout(io.StringIO()):
            result = ai.extract_split_attempt('offline', 0, 'PROMPT', 'DOCUMENT', {}, {})
        self.assertEqual(request.call_count, 2)
        self.assertEqual(result['status'], 'OK')
        self.assertEqual(result['history'][0]['raw_response'], 'broken')
        repair = request.call_args.args[1]
        for text in ('DOCUMENT', 'bad JSON', 'broken'):
            self.assertIn(text, repair)

    def test_invalid_repair_does_not_loop(self):
        with patch.object(ai, 'ask_model_with_deadline', return_value='broken') as request, \
             contextlib.redirect_stdout(io.StringIO()):
            result = ai.extract_split_attempt('offline', 0, '', '', {}, {})
        self.assertEqual(request.call_count, 2)
        self.assertEqual(result['status'], 'INVALID')

    def test_timeout_is_not_repaired(self):
        with patch.object(ai, 'ask_model_with_deadline', side_effect=ai.requests.Timeout) as request:
            result = ai.extract_split_attempt('offline', 0, '', '', {}, {})
        self.assertEqual(request.call_count, 1)
        self.assertEqual(result['status'], 'TIMEOUT')


if __name__ == '__main__':
    unittest.main()
