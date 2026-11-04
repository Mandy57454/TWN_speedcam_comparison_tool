import pandas as pd
import os
import json
import re
import warnings
from collections import Counter


# 動態取得專案路徑
project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

cities = {
    'taipei': '臺北市',
    'TaoYuan': '桃園市',
    'KeeLung': '基隆市',
    'Science_Park': '科學園區',
    'HsinChu_web': '新竹市',
    'MiaoLi': '苗栗縣',
    'TaiChung': '臺中市',
    'YunLin': '雲林縣',
    'ChiaYi': '嘉義縣',
    'NanTou': '南投縣',
    'TaiNan': '臺南市',
    'TaiTung': '臺東縣',
    'HuaLian': '花蓮縣',
    'YiLan': '宜蘭縣',
    'KinMen': '金門縣',
    'Kaohsiung': '高雄市',
    'ChiaYi_sh': '嘉義市',
    'PengHu': '澎湖縣',
    'ChangHua': '彰化縣',
    'newTaipei': '新北市',
    'HsinChu': '新竹縣',
    'PingTung': '屏東縣',
}

example_path = os.path.join(project_path, r"split_data\input_example.xlsx")
OUTPUT_COLUMNS = ['city', 'district', 'road_1', 'road_2', 'note', 'address',
                  'camera_type', 'speed', 'type']


def has_value(value):
    return not pd.isna(value) and str(value).strip() != ''


def text_value(value):
    return str(value).strip() if has_value(value) else ''


def ensure_district(x, remark, z, city):
    # Remarks may contain 區間測速 or 施工區; they are not district names.
    district = text_value(x) or text_value(z)
    if not district or '分局' in district:
        return ''
    if text_value(city) and district.startswith(text_value(city)):
        district = district[len(text_value(city)):]
    if not district:
        return ''
    district = re.split(r'[、,，]', district)[0]
    if city == '臺北市' and district.endswith(('一', '二')):
        district = district[:-1]
    if district.endswith(('鄉', '鎮', '市', '區')) and district not in {'前鎮', '左鎮', '新市', '平鎮'}:
        return district
    return district + '區'


def check_address(city, district, address):
    address_str = text_value(address)
    for prefix in (text_value(city), text_value(district)):
        if prefix and address_str.startswith(prefix):
            address_str = address_str[len(prefix):]
    return address_str


def safe_int(x):
    if not has_value(x):
        return ''
    try:
        value = float(x)
        return int(value) if value.is_integer() else x
    except (ValueError, TypeError):
        return x


def check_camera_type(filename, sheet_name):
    hint = f'{filename} {sheet_name}'.lower()
    if any(word in hint for word in ('mobile', '移動式', '非固定式')):
        return 5
    if 'average' in hint or '區間' in hint:
        return 964
    return 0


def check_banned(ban, remark, distance):
    length = detection_distance(distance)
    if length is not None and length > 0:
        return 964
    return ban if has_value(ban) else text_value(remark)


def detection_distance(value):
    """Return a nonnegative distance, None for empty/unsupported values."""
    if not has_value(value):
        return None
    match = re.fullmatch(r'(?:約)?\s*(\d+(?:\.\d+)?)\s*(?:公尺|公里|米|km|m)?',
                         text_value(value).replace(',', ''), flags=re.IGNORECASE)
    return float(match.group(1)) if match else None


def keyword_counts(values):
    counts = Counter({'超速': 0, '闖紅燈': 0, '測速': 0})
    for value in values:
        text = re.sub(r'\s+', '', text_value(value))
        for word in counts:
            counts[word] += text.count(word)
    return counts


def transform_workbook(dfs, city_name, filename):
    """Transform validated rows; retain rejected rows and provenance in a report."""
    records, expected_text = [], []
    report = {'input_rows': 0, 'ignored_blank_rows': 0, 'rejected_rows': [],
              'accepted_sources': [], 'type_counts': {'mobile': 0, 'average': 0, 'text': 0}}
    for sheet_name, frame in dfs.items():
        if frame.columns.duplicated().any() or '設置地點' not in frame.columns:
            raise ValueError(f'{sheet_name} 欄位重複或缺少設置地點')
        require_id = '編號' in frame.columns and frame['編號'].map(has_value).any()
        hint = check_camera_type(filename, sheet_name)
        for offset, (_, row) in enumerate(frame.iterrows()):
            report['input_rows'] += 1
            raw = row.to_dict()
            if not any(has_value(value) for value in raw.values()):
                report['ignored_blank_rows'] += 1
                continue
            origin = {'sheet': sheet_name, 'excel_row': offset + 2,
                      'device_id': text_value(raw.get('編號', ''))}
            reasons = []
            address = text_value(raw.get('設置地點', ''))
            if not address or re.fullmatch(r'[\d\s\W_]+', address):
                reasons.append('設置地點空白或只有數字/符號')
            elif address in {'設置地點', '地點', '取締地點', '設置地址'} or re.match(
                    r'^(?:備註|說明|合計|總計|小計|註[:：])', address):
                reasons.append('表頭或附註列')
            if require_id and not has_value(raw.get('編號', '')):
                reasons.append('設備編號空白，可能為接續文字或附註')
            distance = raw.get('偵測距離', '')
            if has_value(distance) and detection_distance(distance) is None:
                reasons.append('偵測距離無法辨識')
            district = ensure_district(raw.get('行政區', ''), raw.get('備註', ''),
                                       raw.get('轄區', ''), city_name)
            address = check_address(city_name, district, address)
            if not address:
                reasons.append('移除城市與行政區後沒有地址')
            camera_type = (hint if hint else check_banned(raw.get('取締項目', ''),
                           raw.get('備註', ''), distance))
            if not has_value(camera_type):
                reasons.append('缺少可判定的取締項目或設備類型')
            if reasons:
                report['rejected_rows'].append({**origin, 'reasons': reasons, 'values': raw})
                continue
            kind = 'mobile' if camera_type == 5 else 'average' if camera_type == 964 else 'text'
            report['type_counts'][kind] += 1
            if kind == 'text':
                expected_text.append(raw.get('取締項目', '') if has_value(raw.get('取締項目', ''))
                                     else raw.get('備註', ''))
            records.append({'city': city_name, 'district': district, 'road_1': '', 'road_2': '',
                'note': '', 'address': address, 'camera_type': camera_type,
                'speed': safe_int(raw.get('速限', '')), 'type': ''})
            report['accepted_sources'].append({**origin, 'output_row': len(records) + 1})
    output = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    actual_text = [value for value in output['camera_type'] if value not in (5, 964)]
    before, after = keyword_counts(expected_text), keyword_counts(actual_text)
    report.update(accepted_rows=len(output), rejected_count=len(report['rejected_rows']),
                  input_keywords=dict(before), output_keywords=dict(after))
    if before != after or report['input_rows'] != (
            len(output) + report['rejected_count'] + report['ignored_blank_rows']):
        raise ValueError('筆數或執法關鍵文字核對失敗')
    return output, report


def read_json(path):
    with open(path, encoding='utf-8') as stream:
        report = json.load(stream)
    if not isinstance(report, dict):
        raise ValueError('診斷檔必須是 JSON object')
    return report


def write_json(path, report):
    with open(path, 'w', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2, default=str)


def validate_processing_report(report, dfs):
    if report.get('status') != 'ok' or report.get('issues') or any(
            sheet.get('issues') for sheet in report.get('sheets', [])):
        raise ValueError('上游標示 needs_review 或有未解決問題')
    if report.get('converter', {}).get('status') == 'needs_review':
        raise ValueError('converter 結果仍待檢查')
    if dfs is None:
        return
    counts = report.get('counts', {})
    if 'output_rows' in counts and counts['output_rows'] != sum(len(df) for df in dfs.values()):
        raise ValueError('輸入筆數與診斷檔不一致，可能是過期診斷檔')
    ids = sum(int(df['編號'].map(has_value).sum()) for df in dfs.values() if '編號' in df)
    if 'output_ids' in counts and counts['output_ids'] != ids:
        raise ValueError('輸入設備編號數與診斷檔不一致')
    actual = keyword_counts(value for df in dfs.values()
                            for row in df.itertuples(index=False, name=None) for value in row)
    if any(actual[word] != count for word, count in counts.get('output_keywords', {}).items()):
        raise ValueError('輸入執法文字與診斷檔不一致')
    for sheet in report.get('sheets', []):
        if sheet['sheet'] not in dfs or sheet.get('rows') != len(dfs[sheet['sheet']]):
            raise ValueError('工作表與診斷檔不一致')


def build_city_data(folder_path, city_name):
    """Read-only city assembly; upstream review files never enter formal data."""
    accepted = []
    report = {'city': city_name, 'files': [], 'issues': []}
    if not os.path.isdir(folder_path):
        report['issues'].append('來源資料夾不存在')
        filenames = []
    else:
        filenames = sorted(name for name in os.listdir(folder_path)
                           if name.lower().endswith('.xlsx') and not name.startswith('~$'))
    for filename in filenames:
        file_path = os.path.join(folder_path, filename)
        diagnostic = os.path.splitext(file_path)[0] + '.processing.json'
        item = {'file': filename, 'status': 'skipped', 'processing_report': diagnostic}
        try:
            upstream = read_json(diagnostic) if os.path.exists(diagnostic) else None
            item['upstream_status'] = upstream.get('status') if upstream is not None else 'not_provided'
            if upstream is not None:
                item['declared_rows'] = upstream.get('counts', {}).get('output_rows')
            if upstream is not None:
                validate_processing_report(upstream, None)
            with pd.ExcelFile(file_path) as workbook:
                dfs = pd.read_excel(workbook, sheet_name=None, dtype=object, keep_default_na=False)
            if upstream is not None:
                validate_processing_report(upstream, dfs)
            frame, details = transform_workbook(dfs, city_name, filename)
            # The row numbers refer to the city workbook, not just this input file.
            offset = sum(len(previous) for previous in accepted)
            for origin in details['accepted_sources']:
                origin['output_row'] += offset
            accepted.append(frame)
            item.update(details, status='partial' if details['rejected_count'] else 'ok')
        except Exception as error:
            item['reason'] = str(error)
        report['files'].append(item)
    output = pd.concat(accepted, ignore_index=True) if accepted else pd.DataFrame(columns=OUTPUT_COLUMNS)
    output = output.reindex(columns=OUTPUT_COLUMNS)
    report['output_rows'] = len(output)
    report['ready_for_merge'] = not output.empty
    problems = any(item['status'] != 'ok' for item in report['files'])
    if not filenames:
        report['issues'].append('沒有可處理的 Excel')
    report['status'] = ('needs_review' if output.empty else
                        'partial' if problems or report['issues'] else 'ok')
    return output, report


def process_city(folder_path, city_name, output_path):
    frame, report = build_city_data(folder_path, city_name)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    # Empty output replaces stale results when every source is rejected.
    frame.to_excel(output_path, index=False)
    report_path = os.path.splitext(output_path)[0] + '.assembly.json'
    write_json(report_path, report)
    if report['status'] != 'ok':
        warnings.warn(f'{city_name} 僅輸出已驗證的 {len(frame)} 筆；排除原因詳見 {report_path}', stacklevel=2)
    return report


def combine_excel(folder_path=None, data_combined=None, filenames=None):
    folder_path = folder_path or os.path.join(project_path, 'split_data', 'input', 'all_cities')
    data_combined = data_combined or os.path.join(project_path, 'split_data', 'input', 'all_cities.xlsx')
    # Explicit filenames prevent stale results from previous runs being merged.
    if filenames is None:
        filenames = [f'{name}.xlsx' for name in cities
                     if os.path.isfile(os.path.join(folder_path, f'{name}.xlsx'))]
    frames, report = [], {'files': [], 'status': 'ok'}
    for filename in filenames:
        if not filename.lower().endswith('.xlsx') or filename.lower() == 'mio_data.xlsx':
            continue
        path = os.path.join(folder_path, filename)
        item = {'file': filename, 'status': 'skipped'}
        try:
            diagnostic = os.path.splitext(path)[0] + '.assembly.json'
            city_report = read_json(diagnostic) if os.path.exists(diagnostic) else None
            if city_report is not None and (city_report.get('ready_for_merge') is not True
                    or city_report.get('status') not in ('ok', 'partial')):
                raise ValueError('此城市沒有可合併的已驗證資料')
            frame = pd.read_excel(path, dtype=object, keep_default_na=False)
            if not set(OUTPUT_COLUMNS).issubset(frame.columns):
                raise ValueError('城市輸出缺少必要九欄')
            if city_report is not None and city_report.get('output_rows') != len(frame):
                raise ValueError('城市輸出與診斷筆數不一致')
            frame = frame.reindex(columns=OUTPUT_COLUMNS)
            # split_tool.main uses this existing provenance column to fill type.
            frame['Source_File'] = filename
            frames.append(frame)
            item.update(status='ok', rows=len(frame))
        except Exception as error:
            item['reason'] = str(error)
            report['status'] = 'partial'
        report['files'].append(item)
    columns = OUTPUT_COLUMNS + ['Source_File']
    output = pd.concat(frames, ignore_index=True).reindex(columns=columns) if frames else pd.DataFrame(columns=columns)
    os.makedirs(os.path.dirname(os.path.abspath(data_combined)), exist_ok=True)
    output.to_excel(data_combined, index=False)
    report['output_rows'] = len(output)
    write_json(os.path.splitext(data_combined)[0] + '.assembly.json', report)
    return report


def data_process2(city_map=None, base_path=None):
    root = base_path or project_path
    selected = cities if city_map is None else city_map
    output_folder = os.path.join(root, 'split_data', 'input', 'all_cities')
    reports, filenames = {}, []
    for folder_name, city_name in selected.items():
        output_path = os.path.join(output_folder, f'{folder_name}.xlsx')
        reports[folder_name] = process_city(
            os.path.join(root, 'split_data', 'standard_in', folder_name), city_name, output_path)
        filenames.append(f'{folder_name}.xlsx')
    combined = combine_excel(output_folder,
        os.path.join(root, 'split_data', 'input', 'all_cities.xlsx'), filenames)
    return {'cities': reports, 'combined': combined}


if __name__ == "__main__":
    data_process2()

