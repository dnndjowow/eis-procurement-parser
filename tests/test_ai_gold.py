import contextlib
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest


SPEC = importlib.util.spec_from_file_location('gold_eval', Path(__file__).with_name('evaluate_ai_gold.py'))
gold = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(gold)


class GoldTests(unittest.TestCase):
    @unittest.skipUnless(
        gold.DEFAULT_GOLD.is_file(),
        'Local contract reference is excluded from Git; add tests/gold/contracts_v1.json to run this check.',
    )
    def test_three_block_gold_has_independent_technical_variants(self):
        documents = gold.read_json(gold.DEFAULT_GOLD.read_text())['documents']
        for reference, count in zip(documents, (3, 1)):
            parts = gold.three_block_reference(reference)
            self.assertEqual(set(parts), {'1', '2', '3'})
            self.assertTrue(all('technical_variants' not in f['path'] for f in parts['1']['facts']))
            self.assertEqual(parts['2']['lists'][0]['length'], count)
            self.assertTrue(all(f['path'][0] == 'technical_variants' for f in parts['2']['facts']))

    def expected(self):
        return {'facts': [
            {'path': ['price'], 'value': '76.58', 'kind': 'number', 'sources': ['DOC/P1']},
            {'path': ['unit'], 'value': 'руб./шт.', 'kind': 'unit', 'sources': ['DOC/P1']},
            {'path': ['country'], 'value': None, 'kind': 'text', 'sources': []},
        ], 'lists': [{'path': ['items'], 'length': 2}]}

    def test_format_equivalence_without_rounding(self):
        data = {'price': '76,5800', 'unit': 'руб./штука*', 'country': None, 'items': ['x', 'x']}
        checks = gold.compare_part(data, self.expected())
        self.assertEqual([c['status'] for c in checks], ['CORRECT', 'CORRECT', 'EXPECTED_NULL', 'CORRECT'])
        data['price'] = '76.581'
        self.assertEqual(gold.compare_part(data, self.expected())[0]['status'], 'WRONG')

    def test_absent_is_not_null_and_repetitions_count(self):
        checks = gold.compare_part({'price': '76.58', 'unit': 'руб./шт.', 'items': ['x']}, self.expected())
        self.assertEqual(checks[2]['status'], 'MISSING')
        self.assertEqual(checks[3]['status'], 'MISSING')

    def test_extra_records_visible(self):
        checks = gold.compare_part({'items': ['x', 'x', 'unexpected']}, self.expected())
        self.assertEqual(checks[-1]['status'], 'EXTRA')
        self.assertEqual(checks[-1]['actual'], 'unexpected')

    def test_bad_shape_and_fabricated_country(self):
        checks = gold.compare_part({'price': {'value': '76.58'}, 'country': 'Россия', 'items': {}}, self.expected())
        self.assertEqual(checks[0]['status'], 'WRONG')
        self.assertEqual(checks[2]['status'], 'WRONG')
        self.assertEqual(checks[3]['status'], 'WRONG')

    def test_json_duplicate_rejected(self):
        with self.assertRaises(ValueError):
            gold.read_json('{"x":1,"x":2}')

    def test_fixed_currency_is_not_scored_as_model_extraction(self):
        expected = {'facts': [{'path': ['currency'], 'value': 'руб.',
                              'kind': 'unit', 'sources': ['DOC/P1']}], 'lists': []}
        self.assertEqual(gold.compare_part({}, expected, fixed_currency=True), [])
        self.assertEqual(gold.compare_part({}, expected)[0]['status'], 'MISSING')

    def test_document_hash_and_api_failures(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            folder = root / 'docs'
            folder.mkdir()
            source = folder / 'sample.docx'
            source.write_bytes(b'fixed fixture')
            run = root / 'run'
            item = run / '001'
            item.mkdir(parents=True)
            (run / 'settings.json').write_text(json.dumps({'schema_version': 'split_compact_v2', 'documents': [source.name]}))
            reference = root / 'gold.json'
            reference.write_text(json.dumps({'documents': [{'file': source.name, 'sha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'parts': {'1': self.expected()}}]}))
            (item / 'attempt_01.json').write_text(json.dumps({'model': 'offline', 'part': 1, 'status': 'TIMEOUT', 'seconds': 100}))
            with contextlib.redirect_stdout(io.StringIO()):
                output = gold.evaluate_run(run, folder, reference)
            self.assertEqual(json.loads((output / 'results.json').read_text())[0]['evaluation'], 'NO_RESPONSE')
            source.write_bytes(b'changed fixture')
            with contextlib.redirect_stdout(io.StringIO()):
                output = gold.evaluate_run(run, folder, reference)
            self.assertEqual(json.loads((output / 'results.json').read_text())[0]['reason'], 'document_hash_mismatch')


if __name__ == '__main__':
    unittest.main()
