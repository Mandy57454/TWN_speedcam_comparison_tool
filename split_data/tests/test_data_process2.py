"""Regression tests for formal output, review gates and pipeline provenance."""
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


stage = load_module('stage_two', ROOT / 'split_data/data_process2.py')


def source_frame(rows=None):
    rows = rows if rows is not None else [{
        '編號': 1, '行政區': '西區', '設置地點': '西區甲路', '取締項目': '超速',
        '速限': 50, '偵測距離': ''}]
    return pd.DataFrame(rows)


def write_source(folder, name='fixed.xlsx', frame=None, report=None):
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    (source_frame() if frame is None else frame).to_excel(path, index=False)
    if report is not None:
        path.with_suffix('.processing.json').write_text(json.dumps(report), encoding='utf-8')
    return path


class ClassificationTests(unittest.TestCase):
    def test_empty_and_zero_distances_are_not_average_speed(self):
        for distance in ('', None, float('nan'), 0, '0'):
            self.assertEqual(stage.check_banned('超速', '', distance), '超速')
        for distance in (500, '500公尺', '1.2km'):
            self.assertEqual(stage.check_banned('超速', '', distance), 964)

    def test_type_hints_and_remark_fallback_are_preserved(self):
        self.assertEqual(stage.check_camera_type('mobile1.xlsx', 'Sheet_1'), 5)
        self.assertEqual(stage.check_camera_type('固定設備.xlsx', '非固定式'), 5)
        self.assertEqual(stage.check_camera_type('average.xlsx', 'Sheet_1'), 964)
        self.assertEqual(stage.check_camera_type('高雄市.xlsx', '區間平均速率'), 964)
        self.assertEqual(stage.check_banned('', '闖紅燈、測速', ''), '闖紅燈、測速')

    def test_district_names_and_address_contents_are_not_truncated(self):
        self.assertEqual(stage.ensure_district('二水鄉', '', '', '彰化縣'), '二水鄉')
        self.assertEqual(stage.ensure_district('二林鎮', '', '', '彰化縣'), '二林鎮')
        self.assertEqual(stage.ensure_district('前鎮', '', '', '高雄市'), '前鎮區')
        self.assertEqual(stage.ensure_district('臺中市西區', '', '', '臺中市'), '西區')
        self.assertEqual(stage.ensure_district('', '區間測速', '', '臺中市'), '')
        self.assertEqual(stage.ensure_district('', '', '第一分局', '臺中市'), '')
        self.assertEqual(stage.check_address('臺中市', '西區', '臺中市西區甲路'), '甲路')
        self.assertEqual(stage.check_address('臺中市', '西區', '甲路往西區方向'), '甲路往西區方向')
        self.assertEqual(stage.check_address('臺中市', '', None), '')

    def test_speed_conversion_does_not_truncate(self):
        self.assertEqual(stage.safe_int('50.0'), 50)
        self.assertEqual(stage.safe_int(50.5), 50.5)
        self.assertEqual(stage.safe_int('大車80/小車90'), '大車80/小車90')

    def test_valid_rows_have_exactly_nine_columns_and_intact_types(self):
        frame = source_frame([
            {'編號': 1, '設置地點': '甲路', '取締項目': '闖紅燈、超速', '偵測距離': ''},
            {'編號': 2, '設置地點': '乙路', '取締項目': '超速', '偵測距離': '500公尺'}])
        output, report = stage.transform_workbook({'Sheet_1': frame}, '臺中市', 'fixed.xlsx')
        self.assertEqual(list(output.columns), stage.OUTPUT_COLUMNS)
        self.assertEqual(output['camera_type'].tolist(), ['闖紅燈、超速', 964])
        self.assertEqual(report['input_keywords'], report['output_keywords'])
        self.assertEqual(report['type_counts'], {'mobile': 0, 'average': 1, 'text': 1})

    def test_orphans_headers_empty_addresses_and_notes_are_rejected(self):
        frame = source_frame([
            {'編號': 1, '設置地點': '甲路', '取締項目': '超速'},
            {'編號': None, '設置地點': '接續文字', '取締項目': '超速'},
            {'編號': 2, '設置地點': '', '取締項目': '超速'},
            {'編號': '編號', '設置地點': '設置地點', '取締項目': '取締項目'},
            {'編號': 3, '設置地點': '備註：設備停用', '取締項目': '超速'},
            {'編號': None, '設置地點': None, '取締項目': None}])
        output, report = stage.transform_workbook({'Sheet_1': frame}, '臺中市', 'fixed.xlsx')
        self.assertEqual(len(output), 1)
        self.assertEqual(report['rejected_count'], 4)
        self.assertEqual(report['ignored_blank_rows'], 1)
        self.assertEqual(report['rejected_rows'][0]['excel_row'], 3)
        self.assertEqual(report['accepted_sources'][0]['device_id'], '1')

    def test_no_id_column_is_allowed_for_valid_legacy_equipment(self):
        frame = source_frame([{'設置地點': '甲路', '取締項目': '超速'}])
        output, report = stage.transform_workbook({'Sheet_1': frame}, '臺中市', 'fixed.xlsx')
        self.assertEqual(len(output), 1)
        self.assertEqual(report['rejected_count'], 0)

    def test_unrecognized_distance_and_missing_type_are_retained_in_rejections(self):
        frame = source_frame([
            {'設置地點': '甲路', '取締項目': '超速', '偵測距離': '未知'},
            {'設置地點': '乙路', '取締項目': '', '偵測距離': ''}])
        output, report = stage.transform_workbook({'Sheet_1': frame}, '臺中市', 'fixed.xlsx')
        self.assertTrue(output.empty)
        self.assertEqual(report['rejected_count'], 2)
        self.assertEqual(report['rejected_rows'][0]['values']['偵測距離'], '未知')


class AssemblyTests(unittest.TestCase):
    def test_upstream_needs_review_blocks_file_before_read_and_leaves_it_unchanged(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            path = write_source(folder, report={'status': 'needs_review'})
            original = path.read_bytes()
            with patch.object(stage.pd, 'ExcelFile') as reader:
                output, report = stage.build_city_data(folder, '臺中市')
            reader.assert_not_called()
            self.assertTrue(output.empty)
            self.assertEqual(report['files'][0]['status'], 'skipped')
            self.assertEqual(path.read_bytes(), original)

    def test_invalid_json_does_not_block_other_valid_sources(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            bad = write_source(folder, 'bad.xlsx')
            bad.with_suffix('.processing.json').write_text('invalid JSON', encoding='utf-8')
            write_source(folder, 'fixed.xlsx')
            output, report = stage.build_city_data(folder, '臺中市')
            self.assertEqual(len(output), 1)
            self.assertEqual(report['status'], 'partial')
            self.assertEqual(report['files'][0]['status'], 'skipped')
            self.assertEqual(report['files'][1]['upstream_status'], 'not_provided')

    def test_stale_row_and_keyword_diagnostics_block_sources(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            write_source(folder, 'count.xlsx', report={'status': 'ok', 'counts': {'output_rows': 2}})
            write_source(folder, 'words.xlsx', report={
                'status': 'ok', 'counts': {'output_rows': 1, 'output_keywords': {'超速': 0}}})
            output, report = stage.build_city_data(folder, '臺中市')
            self.assertTrue(output.empty)
            self.assertTrue(all(item['status'] == 'skipped' for item in report['files']))

    def test_partial_city_contains_only_valid_rows_and_can_merge(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'sources'
            write_source(source, 'review.xlsx', report={'status': 'needs_review'})
            good = write_source(source)
            original = good.read_bytes()
            output = root / 'TaiChung.xlsx'
            with self.assertWarns(UserWarning):
                report = stage.process_city(source, '臺中市', output)
            self.assertEqual(report['status'], 'partial')
            self.assertTrue(report['ready_for_merge'])
            self.assertEqual(list(pd.read_excel(output).columns), stage.OUTPUT_COLUMNS)
            self.assertEqual(good.read_bytes(), original)
            combined = root / 'combined.xlsx'
            result = stage.combine_excel(root, combined, ['TaiChung.xlsx'])
            self.assertEqual(result['output_rows'], 1)
            self.assertEqual(pd.read_excel(combined)['Source_File'].tolist(), ['TaiChung.xlsx'])

    def test_empty_current_run_replaces_stale_city_and_combined_results(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'split_data/standard_in/TaiChung'
            write_source(source, report={'status': 'needs_review'})
            target = root / 'split_data/input/all_cities'
            target.mkdir(parents=True)
            stale = pd.DataFrame([dict.fromkeys(stage.OUTPUT_COLUMNS, 'old')])
            stale.to_excel(target / 'TaiChung.xlsx', index=False)
            stale.to_excel(target / 'HsinChu.xlsx', index=False)
            with self.assertWarns(UserWarning):
                result = stage.data_process2({'TaiChung': '臺中市'}, root)
            self.assertTrue(pd.read_excel(target / 'TaiChung.xlsx').empty)
            combined = root / 'split_data/input/all_cities.xlsx'
            self.assertTrue(pd.read_excel(combined).empty)
            self.assertEqual(result['combined']['output_rows'], 0)
            self.assertEqual([item['file'] for item in result['combined']['files']], ['TaiChung.xlsx'])

    def test_combine_only_requested_city_workbooks_and_preserve_source_column(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            frame = pd.DataFrame([dict.fromkeys(stage.OUTPUT_COLUMNS, '')])
            frame['city'], frame['address'], frame['camera_type'] = '臺中市', '甲路', '超速'
            for name in ('TaiChung.xlsx', 'HsinChu.xlsx', 'mio_data.xlsx'):
                frame.to_excel(root / name, index=False)
            (root / 'diagnostic.json').write_text('{}')
            combined = root / 'merged.xlsx'
            stage.combine_excel(root, combined, ['TaiChung.xlsx', 'mio_data.xlsx', 'diagnostic.json'])
            result = pd.read_excel(combined)
            self.assertEqual(list(result.columns), stage.OUTPUT_COLUMNS + ['Source_File'])
            self.assertEqual(len(result), 1)
            self.assertEqual(result.iloc[0]['Source_File'], 'TaiChung.xlsx')

    def test_empty_combine_has_a_stable_schema(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'empty.xlsx'
            report = stage.combine_excel(directory, path, [])
            self.assertEqual(report['output_rows'], 0)
            self.assertEqual(list(pd.read_excel(path).columns), stage.OUTPUT_COLUMNS + ['Source_File'])


class FullPipelineTests(unittest.TestCase):
    def test_pdf_to_formal_city_and_combined_data_preserves_speed_items(self):
        pdf = ROOT / 'selenium/city_data/TaiChung/臺中市政府警察局執行固定式科學儀器執法設備取締地點一覽表.pdf'
        if not pdf.exists():
            self.skipTest('Local reference PDF is not available')
        converter = load_module('converter_for_stage_two', ROOT / 'converter/pdf_to_excel.py')
        first = load_module('stage_one_for_stage_two', ROOT / 'split_data/data_process1.py')
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            intermediate = root / 'source.xlsx'
            source = root / 'split_data/standard_in/TaiChung'
            source.mkdir(parents=True)
            converter.convert_pdf(pdf, intermediate)
            first.process_workbook(intermediate, source / 'fixed.xlsx')
            reports = stage.data_process2({'TaiChung': '臺中市'}, root)
            output = pd.read_excel(root / 'split_data/input/all_cities/TaiChung.xlsx')
            combined = pd.read_excel(root / 'split_data/input/all_cities.xlsx')
            self.assertEqual(list(output.columns), stage.OUTPUT_COLUMNS)
            self.assertEqual(len(output), 227)
            self.assertEqual(int(output['camera_type'].str.contains('超速').sum()), 174)
            self.assertFalse((output['camera_type'] == 964).any())
            self.assertEqual(len(combined), 227)
            self.assertTrue((combined['Source_File'] == 'TaiChung.xlsx').all())
            city_report = reports['cities']['TaiChung']
            self.assertEqual(city_report['status'], 'ok')
            self.assertEqual(city_report['files'][0]['accepted_rows'], 227)
            self.assertEqual(city_report['files'][0]['input_keywords'], city_report['files'][0]['output_keywords'])
            saved = json.loads((root / 'split_data/input/all_cities/TaiChung.assembly.json').read_text(encoding='utf-8'))
            self.assertEqual(saved['files'][0]['accepted_sources'][0]['output_row'], 2)


if __name__ == '__main__':
    unittest.main()
