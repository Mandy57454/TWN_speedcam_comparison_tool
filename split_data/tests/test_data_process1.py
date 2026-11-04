"""Run from the repository root: python -m unittest discover -s split_data/tests -v."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


split = load_module('split_process', ROOT / 'split_data/data_process1.py')


class StandardizationTests(unittest.TestCase):
    def test_missing_region_does_not_delete_equipment(self):
        frame = pd.DataFrame([[1, None, '甲路', '超速'], [2, '西區', '乙路', '闖紅燈']],
                             columns=['編號', '行政區', '設置地點', '取締項目'])
        sheets, report = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(sheets['Sheet_1']['編號'].tolist(), ['1', '2'])
        self.assertEqual(report['counts']['output_keywords']['超速'], 1)
        self.assertEqual(report['status'], 'needs_review')

    def test_column_order_and_first_two_blanks_do_not_filter_rows(self):
        frame = pd.DataFrame([[None, None, 7, '甲路', '超速']],
                             columns=['備註', '其他', '設備編號', '地點', '執法項目'])
        sheets, report = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(sheets['Sheet_1'].iloc[0]['編號'], '7')
        self.assertEqual(report['status'], 'ok')

    def test_unknown_format_keeps_more_than_ten_rows(self):
        frame = pd.DataFrame([[f'原始資料{i}', '超速'] for i in range(15)])
        sheets, report = split.standardize_sheets({'unknown': frame})
        self.assertEqual(len(sheets['unknown']), 15)
        self.assertEqual(report['counts']['output_keywords']['超速'], 15)
        self.assertFalse(report['sheets'][0]['recognized'])
        self.assertEqual(report['status'], 'needs_review')

    def test_data_values_are_not_replaced_with_column_aliases(self):
        frame = pd.DataFrame([[1, '功能', '方向']], columns=['項次', '地點', '功能'])
        sheets, _ = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(sheets['Sheet_1'].iloc[0]['設置地點'], '功能')
        self.assertEqual(sheets['Sheet_1'].iloc[0]['取締項目'], '方向')

    def test_repeated_alias_headers_removed_and_new_schema_split(self):
        frame = pd.DataFrame([
            ['設備編號', '地點', '區域'], ['1', '甲路', '西區'],
            ['項次', '設置地點', '行政區'], ['2', '乙路', '北區'],
            ['編號', '地點', '速限'], ['3', '丙路', '50']])
        sheets, report = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(list(sheets), ['Sheet_1', 'sheet1'])
        self.assertEqual(sheets['Sheet_1']['編號'].tolist(), ['1', '2'])
        self.assertEqual(sheets['sheet1'].iloc[0]['速限'], '50')
        self.assertEqual(report['counts']['input_ids'], 3)
        self.assertEqual(report['counts']['output_ids'], 3)

    def test_lone_bianhao_in_a_note_is_not_a_header(self):
        frame = pd.DataFrame([[1, '甲路', '編號']], columns=['編號', '設置地點', '備註'])
        sheets, _ = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(len(sheets), 1)
        self.assertEqual(sheets['Sheet_1'].iloc[0]['備註'], '編號')

    def test_header_after_long_preamble_and_two_line_header(self):
        rows = [[f'標題{i}', None, None] for i in range(12)]
        rows += [['設備', '行政區', '地點'], ['編號', None, None], ['1', '西區', '甲路']]
        sheets, report = split.standardize_sheets({'Sheet_1': pd.DataFrame(rows)})
        self.assertEqual(sheets['Sheet_1']['編號'].tolist(), ['1'])
        self.assertEqual(len(report['sheets'][0]['preamble']), 12)
        self.assertEqual(report['sheets'][0]['source_rows'], [15])

    def test_duplicate_columns_preserve_values_and_original_suffixes(self):
        frame = pd.DataFrame([[1, '甲路', '50', '60', '70']],
                             columns=['編號', '地點', '速限', '速限', '速限1'])
        sheets, report = split.standardize_sheets({'Sheet_1': frame})
        result = sheets['Sheet_1']
        self.assertTrue(result.columns.is_unique)
        self.assertEqual(result.iloc[0][['速限', '速限2', '速限1']].tolist(), ['50', '60', '70'])
        self.assertEqual(report['status'], 'needs_review')

    def test_orphan_text_and_partial_rows_are_not_dropped_or_guessed(self):
        frame = pd.DataFrame([[1, '甲路', '闖紅燈'], [None, None, '超速'],
                              [None, None, None]], columns=['編號', '設置地點', '取締項目'])
        sheets, report = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(len(sheets['Sheet_1']), 2)
        self.assertEqual(sheets['Sheet_1'].iloc[1]['取締項目'], '超速')
        self.assertEqual(report['counts']['input_rows'], report['counts']['output_rows'])
        self.assertEqual(report['status'], 'needs_review')

    def test_converter_review_status_is_propagated(self):
        frame = pd.DataFrame([[1, '甲路']], columns=['編號', '設置地點'])
        upstream = {'status': 'needs_review', 'issues': ['來源未解析']}
        sheets, report = split.standardize_sheets({'Sheet_1': frame}, upstream)
        self.assertEqual(len(sheets['Sheet_1']), 1)
        self.assertEqual(report['converter'], upstream)
        self.assertEqual(report['status'], 'needs_review')

    def test_stale_converter_count_is_reported(self):
        frame = pd.DataFrame([[1, '甲路']], columns=['編號', '設置地點'])
        _, report = split.standardize_sheets({'Sheet_1': frame},
            {'status': 'ok', 'device_count': {'output': 2}})
        self.assertTrue(any('版本' in issue for issue in report['issues']))

    def test_accidental_data_loss_is_detected(self):
        frame = pd.DataFrame([[1, '甲路'], [2, '乙路']], columns=['編號', '設置地點'])
        with patch.object(split, 'reorder_columns', side_effect=lambda frame: frame.iloc[:1]):
            _, report = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(report['missing_ids'], ['2'])
        self.assertEqual(report['status'], 'needs_review')

    def test_parse_exception_preserves_whole_sheet(self):
        frame = pd.DataFrame([['未知', '超速'], ['接續', '內容']])
        with patch.object(split, 'parse_sheet', side_effect=ValueError('test failure')):
            sheets, report = split.standardize_sheets({'Sheet_1': frame})
        self.assertEqual(len(sheets['Sheet_1']), 2)
        self.assertEqual(report['counts']['output_keywords']['超速'], 1)
        self.assertEqual(report['status'], 'needs_review')

    def test_long_case_insensitive_sheet_names_do_not_collide(self):
        used = {'A' * 31}
        name = split.safe_sheet_name('a' * 40 + '[]:*?', used)
        self.assertLessEqual(len(name), 31)
        self.assertNotEqual(name.casefold(), ('A' * 31).casefold())
        self.assertFalse(any(char in name for char in '[]:*?/\\'))

    def test_file_export_and_invalid_converter_json(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / '來源.xlsx'
            output = Path(directory) / 'fixed.xlsx'
            pd.DataFrame([[1, None, '甲路', '超速']],
                columns=['編號', '行政區', '地點', '取締項目']).to_excel(source, index=False)
            source.with_suffix('.conversion.json').write_text('invalid JSON', encoding='utf-8')
            with self.assertWarns(UserWarning):
                report = split.process_workbook(source, output)
            result = pd.read_excel(output)
            self.assertEqual(result.iloc[0]['取締項目'], '超速')
            self.assertEqual(report['status'], 'needs_review')
            saved = json.loads(output.with_suffix('.processing.json').read_text(encoding='utf-8'))
            self.assertIn('診斷檔讀取失敗', saved['converter']['issues'][0])
            self.assertEqual(list(result.columns[:7]),
                ['編號', '行政區', '設置地點', '速限', '取締項目', '拍攝方向', '偵測距離'])


class PdfPipelineTests(unittest.TestCase):
    def test_taichung_pdf_converter_split_roundtrip(self):
        pdf = ROOT / 'selenium/city_data/TaiChung/臺中市政府警察局執行固定式科學儀器執法設備取締地點一覽表.pdf'
        if not pdf.exists():
            self.skipTest('Local reference PDF is not available')
        converter = load_module('pdf_converter', ROOT / 'converter/pdf_to_excel.py')
        with tempfile.TemporaryDirectory() as directory:
            intermediate = Path(directory) / (pdf.stem + '.xlsx')
            output = Path(directory) / 'fixed.xlsx'
            converter_report = converter.convert_pdf(pdf, intermediate)
            report = split.process_workbook(intermediate, output)
            result = pd.read_excel(output)
            self.assertEqual(converter_report['status'], 'ok')
            self.assertEqual(report['status'], 'ok')
            self.assertEqual(report['counts']['input_ids'], 227)
            self.assertEqual(report['counts']['output_ids'], 227)
            self.assertEqual(result['編號'].astype(int).tolist(), list(range(1, 228)))
            self.assertEqual(int(result['取締項目'].str.contains('超速').sum()), 174)
            self.assertFalse(result['行政區'].isna().any())
            self.assertFalse(result['設置地點'].isna().any())
            self.assertEqual(report['counts']['input_keywords'], report['counts']['output_keywords'])


if __name__ == '__main__':
    unittest.main()
