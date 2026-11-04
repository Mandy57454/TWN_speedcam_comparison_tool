"""Run: python -m unittest discover -s split_data/tests -p test_data_process3.py -v."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


stage = load_module('stage_three', ROOT / 'split_data/data_process3.py')


def final_frame(camera_types, speeds):
    frame = pd.DataFrame([dict.fromkeys(stage.REQUIRED_COLUMNS, '') for _ in camera_types])
    frame['city'], frame['district'], frame['address'] = '臺中市', '西區', '甲路'
    frame['camera_type'], frame['speed'] = camera_types, speeds
    frame['Source_File'] = 'TaiChung.xlsx'
    return frame


class CameraClassificationTests(unittest.TestCase):
    def test_all_numeric_codes_are_exact_and_idempotent(self):
        for token in stage.c_type:
            expected = int(token)
            for value in (expected, float(expected), token, token + '.0'):
                self.assertEqual(stage.get_c_type_value(value), expected)
        self.assertEqual(stage.get_c_type_value('６１'), 61)

    def test_unknown_numbers_do_not_match_a_code_substring(self):
        for value in (611, '9641', '61.5', '50', '1.5e1', True):
            self.assertEqual(stage.get_c_type_value(value), 0)
            self.assertEqual(stage.classify_camera_type(value)['status'], 'unknown')

    def test_digits_in_descriptions_do_not_override_text_keywords(self):
        self.assertEqual(stage.get_c_type_value('違規停車(6-20時)'), 61)
        self.assertEqual(stage.get_c_type_value('臨時停車(7-19時)'), 61)
        self.assertEqual(stage.get_c_type_value('大貨車7-9時禁行'), 99)
        self.assertEqual(stage.get_c_type_value('設備50型'), 0)

    def test_user_selected_original_text_priority_is_preserved(self):
        cases = {'區間、科技執法、超速': 964, '超速、闖紅燈、科技執法': 6,
                 '闖紅燈兼超速': 1, '闖紅燈、違規停車': 3,
                 '違規停車、機車': 61, '機車行駛禁行路段、大貨車': 7,
                 '移動式測速': 1}
        for text, expected in cases.items():
            self.assertEqual(stage.get_c_type_value(text), expected)

    def test_multiple_matches_are_reported_without_changing_priority(self):
        frame = final_frame(['超速、闖紅燈、不依標誌'], [50])
        output, report = stage.classify_dataframe(frame)
        self.assertEqual(output.iloc[0]['camera_type'], 6)
        self.assertEqual(report['status'], 'ok')
        self.assertEqual(report['priority'], [964, 6, 1, 3, 61, 7, 99, 5])
        self.assertEqual([item['code'] for item in report['ambiguous_rows'][0]['matches']], [6, 1, 3])

    def test_missing_and_unknown_types_remain_traceable_as_zero(self):
        frame = final_frame([None, '', '未支援設備', 0], [0, 0, 0, 0])
        output, report = stage.classify_dataframe(frame)
        self.assertEqual(output['camera_type'].tolist(), [0, 0, 0, 0])
        self.assertEqual(len(report['unclassified_rows']), 4)
        self.assertEqual(report['unclassified_rows'][2]['original'], '未支援設備')
        self.assertEqual(report['status'], 'needs_review')

    def test_unicode_and_whitespace_are_normalized(self):
        self.assertEqual(stage.get_c_type_value('超\n 速'), 1)
        self.assertEqual(stage.get_c_type_value('車輛未禮讓行人'), 6)


class SpeedTests(unittest.TestCase):
    def test_only_complete_markers_become_zero(self):
        for marker in ('-', '\\', 'x', 'X', '違左', ' － '):
            self.assertEqual(stage.get_s_type_value(marker), 0)
        for value in ('50-60', '90\\80', 'max50', '違左轉40'):
            self.assertEqual(stage.get_s_type_value(value), value)
            self.assertEqual(stage.classify_speed(value)['status'], 'unresolved')

    def test_scalar_units_and_numbers_are_normalized_without_truncation(self):
        for value in (50, '50.0', '50公里', '50公里/小時', '50 km/h', '５０公里／小時'):
            self.assertEqual(stage.get_s_type_value(value), 50)
        self.assertEqual(stage.get_s_type_value(50.5), 50.5)
        self.assertEqual(stage.get_s_type_value('50.5'), 50.5)

    def test_composite_limits_are_preserved_and_flagged(self):
        values = ['大車80/小車90', '60/40', '50-60', '未知', -1]
        output, report = stage.classify_dataframe(final_frame([5] * len(values), values))
        self.assertEqual(output['speed'].tolist(), values)
        self.assertEqual(len(report['unresolved_speeds']), len(values))
        self.assertEqual(report['status'], 'needs_review')

    def test_empty_speed_is_consistently_zero(self):
        for value in (None, float('nan'), '', 'N/A'):
            self.assertEqual(stage.get_s_type_value(value), 0)


class FrameAndFileTests(unittest.TestCase):
    def test_row_order_other_columns_and_original_frame_are_preserved(self):
        frame = final_frame(['61', '超速'], ['50公里', 60])
        frame.index = [7, 2]
        frame['extra'] = ['原始資料', '另筆']
        original = frame.copy(deep=True)
        output, report = stage.classify_dataframe(frame)
        self.assertTrue(frame.equals(original))
        self.assertEqual(list(output.index), [7, 2])
        self.assertEqual(list(output.columns), list(frame.columns))
        self.assertEqual(output['extra'].tolist(), ['原始資料', '另筆'])
        self.assertEqual(report['input_rows'], report['output_rows'])
        second, _ = stage.classify_dataframe(output)
        self.assertTrue(second.equals(output))

    def test_missing_or_duplicate_classification_columns_raise(self):
        with self.assertRaises(ValueError):
            stage.classify_dataframe(pd.DataFrame({'camera_type': ['超速']}))
        with self.assertRaises(ValueError):
            stage.classify_dataframe(pd.DataFrame([['超速', 50, 60]],
                                    columns=['camera_type', 'speed', 'speed']))

    def test_export_preserves_schema_and_records_original_values(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output = root / 'source.xlsx', root / 'gov_data.xlsx'
            final_frame(['超速、闖紅燈、不依標誌', 61], ['50公里', '50-60']).to_excel(
                source, sheet_name='自訂合併表', index=False)
            original = source.read_bytes()
            with self.assertWarns(UserWarning):
                report = stage.data_process3(source, output)
            result = pd.read_excel(output, keep_default_na=False)
            self.assertEqual(result['camera_type'].tolist(), [6, 61])
            self.assertEqual(result['speed'].tolist(), [50, '50-60'])
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(list(result.columns), stage.REQUIRED_COLUMNS + ['Source_File'])
            saved = json.loads(output.with_suffix('.classification.json').read_text(encoding='utf-8'))
            self.assertEqual(saved['unresolved_speeds'][0]['original'], '50-60')
            self.assertEqual(report['output_rows'], 2)

    def test_invalid_schema_does_not_overwrite_existing_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output = root / 'bad.xlsx', root / 'gov_data.xlsx'
            pd.DataFrame({'camera_type': ['超速'], 'speed': [50]}).to_excel(source, index=False)
            output.write_bytes(b'existing result')
            with self.assertRaises(ValueError):
                stage.data_process3(source, output)
            self.assertEqual(output.read_bytes(), b'existing result')

    def test_multiple_sheets_do_not_silently_drop_data(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output = root / 'multi.xlsx', root / 'gov_data.xlsx'
            with pd.ExcelWriter(source) as writer:
                final_frame(['超速'], [50]).to_excel(writer, sheet_name='Sheet1', index=False)
                final_frame(['闖紅燈'], [50]).to_excel(writer, sheet_name='Sheet2', index=False)
            with self.assertRaises(ValueError):
                stage.data_process3(source, output)
            self.assertFalse(output.exists())

    def test_empty_input_replaces_stale_output_with_empty_schema(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output = root / 'empty.xlsx', root / 'gov_data.xlsx'
            pd.DataFrame(columns=stage.REQUIRED_COLUMNS).to_excel(source, index=False)
            output.write_bytes(b'stale result')
            with self.assertWarns(UserWarning):
                report = stage.data_process3(source, output)
            self.assertTrue(pd.read_excel(output).empty)
            self.assertEqual(report['output_rows'], 0)


class IntegrationTests(unittest.TestCase):
    def test_stage_two_outputs_classify_without_schema_or_row_changes(self):
        previous = load_module('previous_for_stage_three', ROOT / 'split_data/data_process2.py')
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'split_data/standard_in/TaiChung'
            source.mkdir(parents=True)
            fixed = pd.DataFrame([
                {'編號': i + 1, '設置地點': f'西區甲路{i + 1}號', '取締項目': item,
                 '速限': '50-60' if i == 3 else 50, '偵測距離': ''}
                for i, item in enumerate(['超速', '闖紅燈', '超速、不依標誌', '違規停車(6-20時)', '科技執法'])])
            fixed.to_excel(source / 'fixed.xlsx', index=False)
            for name in ('mobile.xlsx', 'average.xlsx'):
                pd.DataFrame([{'設置地點': name + '甲路', '速限': 50}]).to_excel(source / name, index=False)
            previous.data_process2({'TaiChung': '臺中市'}, root)
            combined = root / 'split_data/input/all_cities.xlsx'
            output = root / 'gov_data.xlsx'
            with self.assertWarns(UserWarning):
                report = stage.data_process3(combined, output)
            result = pd.read_excel(output)
            self.assertEqual(len(result), 7)
            self.assertEqual(report['camera_type_counts'],
                             {'1': 1, '3': 1, '5': 1, '6': 2, '61': 1, '964': 1})
            self.assertEqual(report['unclassified_rows'], [])
            self.assertEqual(result['Source_File'].tolist(), ['TaiChung.xlsx'] * 7)
            self.assertEqual(len(report['unresolved_speeds']), 1)


if __name__ == '__main__':
    unittest.main()
