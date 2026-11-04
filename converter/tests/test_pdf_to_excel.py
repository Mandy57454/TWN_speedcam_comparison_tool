"""Regression checks for schema detection, merged cells and PDF completeness.

Run: python -m unittest discover -s converter/tests -v
"""
import importlib.util
from pathlib import Path
from types import SimpleNamespace
import json
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location('pdf_converter', ROOT / 'converter/pdf_to_excel.py')
converter = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(converter)


class ParsingTests(unittest.TestCase):
    def test_multiline_header_does_not_consume_first_record(self):
        rows = [['title', None, None], ['設備', '行政區', '設置地點'],
                ['編號', None, None], ['1', '西區', '五權路']]
        frame, = converter.convert_to_dataframe([rows])
        self.assertEqual(list(frame.columns), ['設備編號', '行政區', '設置地點'])
        self.assertEqual(frame.values.tolist(), [[1, '西區', '五權路']])

    def test_same_width_different_schemas_stay_separate(self):
        tables = [[['編號', '設置地點', '行政區'], ['1', '甲路', '甲區']],
                  [['編號', '設置地點', '速限'], ['2', '乙路', '50']]]
        frames = converter.convert_to_dataframe(tables)
        self.assertEqual(len(frames), 2)
        self.assertEqual(list(frames[1].columns), ['編號', '設置地點', '速限'])

    def test_repeated_headers_removed_and_same_schema_pages_combined(self):
        header = ['編號', '設置地點', '行政區']
        tables = [[header, ['1', '甲路', '甲區'], header, ['2', '乙路', '乙區']],
                  [header, ['3', '丙路', '丙區']]]
        frame, = converter.convert_to_dataframe(tables)
        self.assertEqual(frame['編號'].tolist(), [1, 2, 3])

    def test_unknown_tables_keep_all_values_without_borrowed_header(self):
        frames = converter.convert_to_dataframe([
            [['編號', '設置地點', '行政區'], ['1', '甲路', '甲區']],
            [['未知標題', None, None], ['A', 'B', 'C']]])
        self.assertFalse(frames[1].attrs['recognized'])
        self.assertEqual(frames[1].iloc[1].tolist(), ['A', 'B', 'C'])

    def test_vertical_merge_filled_but_real_blank_left_blank(self):
        header = [(0, 0, 10, 10), (10, 0, 20, 10), (20, 0, 40, 10)]
        cells = [header,
                 [(0, 10, 10, 20), (10, 10, 20, 30), (20, 10, 40, 20)],
                 [(0, 20, 10, 30), None, (20, 20, 40, 30)],
                 [(0, 30, 10, 40), (10, 30, 20, 40), (20, 30, 40, 40)]]
        table = SimpleNamespace(rows=[SimpleNamespace(cells=row) for row in cells])
        values = [['編號', '行政區', '設置地點'], ['1', '西區', '甲路'],
                  ['2', None, '乙路'], ['3', '', '丙路']]
        result = converter.fill_vertical_spans(table, values)
        self.assertEqual(result[2][1], '西區')
        self.assertEqual(result[3][1], '')
        self.assertIsNone(values[2][1])

    def test_fragments_are_retained_and_flagged(self):
        tables = [[['編號', '設置地點', '取締項目'],
                   ['1', '甲路', '闖紅燈'], [None, None, '超速']]]
        frames = converter.convert_to_dataframe(tables)
        report = converter.conversion_report(tables, frames)
        self.assertEqual(len(frames[0]), 2)
        self.assertEqual(report['status'], 'needs_review')
        self.assertIn('超速', frames[0].iloc[1, 2])

    def test_blank_coordinates_and_out_of_range_values_are_flagged(self):
        tables = [[['編號', '設置地點', '座標緯度'],
                   ['1', '甲路', '100'], ['2', '乙路', None]]]
        report = converter.conversion_report(tables, converter.convert_to_dataframe(tables))
        self.assertTrue(any('座標緯度' in issue for issue in report['sheets'][0]['issues']))

    def test_different_section_within_one_table_keeps_its_header(self):
        rows = [['編號', '設置地點', '行政區'], ['1', '甲路', '甲區'],
                ['編號', '設置地點', '速限'], ['2', '乙路', '60']]
        frames = converter.convert_to_dataframe([rows])
        self.assertEqual(len(frames), 2)
        self.assertEqual(frames[1].iloc[0].tolist(), [2, '乙路', 60])

    def test_header_inheritance_requires_geometry_and_number_continuity(self):
        rows = [['2', '西區', '甲路']]
        table = SimpleNamespace(
            extract=lambda: rows,
            rows=[SimpleNamespace(cells=[(0, 0, 10, 10), (10, 0, 20, 10),
                                        (20, 0, 40, 10)])])
        labels = ['編號', '行政區', '設置地點']
        previous = {'labels': labels, 'columns': [(0, .25), (.25, .5), (.5, 1)],
                    'last_id': 1}
        self.assertEqual(converter.candidate_rows([table], 40, previous)[0][0], labels)
        previous['last_id'] = 5
        self.assertEqual(converter.candidate_rows([table], 40, previous)[0], rows)
        previous['last_id'] = 1
        previous['columns'] = [(0, .5), (.5, .75), (.75, 1)]
        self.assertEqual(converter.candidate_rows([table], 40, previous)[0], rows)

    def test_export_preserves_sheet_naming_and_writes_diagnostics(self):
        tables = converter.ExtractedTables()
        tables.append(converter.ExtractedTable(
            [['編號', '行政區', '設置地點'], ['1', '西區', '甲路']],
            1, (0, 0, 100, 100), 'lines'))
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / '原始名稱.xlsx'
            with patch.object(converter, 'extract_tables_from_pdf', return_value=tables):
                report = converter.convert_pdf('原始名稱.pdf', output)
            with converter.pd.ExcelFile(output) as workbook:
                self.assertEqual(workbook.sheet_names, ['Sheet_1'])
                self.assertEqual(converter.pd.read_excel(workbook).iloc[0].tolist(),
                                 [1, '西區', '甲路'])
            saved = json.loads(output.with_suffix('.conversion.json').read_text(encoding='utf-8'))
            self.assertEqual(saved['status'], 'ok')
            self.assertEqual(report['tables'][0]['page'], 1)

    def test_device_loss_is_reported(self):
        tables = [[['編號', '行政區', '設置地點'],
                   ['1', '西區', '甲路'], ['2', '西區', '乙路']]]
        frame, = converter.convert_to_dataframe(tables)
        report = converter.conversion_report(tables, [frame.iloc[:1]])
        self.assertEqual(report['device_count'], {'extracted': 2, 'output': 1})
        self.assertEqual(report['status'], 'needs_review')


class SourcePdfTests(unittest.TestCase):
    def test_taichung_preserves_devices_and_enforcement_items(self):
        source = ROOT / 'selenium/city_data/TaiChung/臺中市政府警察局執行固定式科學儀器執法設備取締地點一覽表.pdf'
        if not source.exists():
            self.skipTest('Local reference PDF is not available')
        tables = converter.extract_tables_from_pdf(source)
        frames = converter.convert_to_dataframe(tables)
        self.assertEqual(len(frames), 1)
        frame = frames[0]
        self.assertEqual(frame['設備編號'].tolist(), list(range(1, 228)))
        self.assertEqual(int(frame['取締項目'].str.contains('超速').sum()), 174)
        self.assertFalse(frame['行政區'].isna().any())
        self.assertFalse(frame['設置地點'].isna().any())
        self.assertIn('超速', frame.loc[frame['設備編號'] == 7, '取締項目'].iloc[0])
        self.assertIn('執法', frame.loc[frame['設備編號'] == 7, '備註'].iloc[0])
        self.assertEqual(converter.conversion_report(tables, frames)['status'], 'ok')


if __name__ == '__main__':
    unittest.main()
