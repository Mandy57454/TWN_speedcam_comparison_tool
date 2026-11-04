import pandas as pd
import os
import json
import re
import unicodedata
import warnings
from collections import Counter
from decimal import Decimal, InvalidOperation


# 動態取得專案路徑
project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

c_type = {
    '964': {'964', '區間', '測速跨線未保安距'},
    '6': {'6', '科技執法', '依標誌', '行人穿越道', '駛入內車道', '駛出外車道', '使用方向燈', '讓行人', '淨空', '違規轉彎', '違左',
          '迴轉', '違規左右', '右轉', '號誌', '逆向', '雙白', '直行車', '不依規定駛入來車道', '違規左轉', '依規定轉彎', '跨越雙黃線', '劃分島'},
    '1': {'1', '固定式', '超速', '測速', '雷達測速', '分局'},
    '3': {'3', '闖紅燈'},
    '61': {'61', '違規停車', '違停', '臨時'},
    '7': {'7', '機車', '行駛人行道', '車輛行駛禁行路段', '禁行車種'},
    '99': {'99', '大貨車', '大型車', '砂石車', '時段性', '取締項目', '行駛禁行路段', '偵測註銷', '停車時間'},
    '5': {'5', '移動式'},
}
s_type = {
    '0': {'-', '\\', 'x', '違左'},
}
REQUIRED_COLUMNS = ['city', 'district', 'road_1', 'road_2', 'note', 'address',
                    'camera_type', 'speed', 'type']


def normalized_text(value):
    if pd.isna(value):
        return ''
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value))).casefold()


def is_missing(value):
    return normalized_text(value) in {'', 'nan', 'none', 'null', 'n/a'}


def decimal_value(text):
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def classify_camera_type(cell):
    """Respect configured dictionary order for text; numeric codes are exact.

    Numeric aliases such as '6' must never match times, speeds or code '61'.
    Multiple text matches remain auditable without changing the user's priority.
    """
    result = {'code': 0, 'status': 'unknown', 'matches': []}
    if is_missing(cell):
        result['status'] = 'missing'
        return result
    text = normalized_text(cell)
    number = decimal_value(text)
    if number is not None:
        if number.is_finite() and number == number.to_integral_value():
            code = str(int(number))
            if code in c_type:
                result.update(code=int(code), status='numeric',
                              matches=[{'code': int(code), 'keywords': [code]}])
        return result
    for code, keywords in c_type.items():
        matched = sorted(keyword for keyword in keywords
                         if not str(keyword).isdigit() and normalized_text(keyword) in text)
        if matched:
            result['matches'].append({'code': int(code), 'keywords': matched})
    if result['matches']:
        result.update(code=result['matches'][0]['code'], status='text')
    return result


def classify_speed(cell):
    """Normalize scalar limits; retain ambiguous ranges instead of guessing."""
    if is_missing(cell):
        return {'value': 0, 'status': 'missing'}
    text = normalized_text(cell)
    for value, markers in s_type.items():
        if text in {normalized_text(marker) for marker in markers}:
            return {'value': int(value), 'status': 'marker'}
    match = re.fullmatch(
        r'(\d+(?:\.\d+)?)(?:公里/小時|公里/時|公里|km/h|kmh|kph)?', text)
    if match:
        number = Decimal(match.group(1))
        value = int(number) if number == number.to_integral_value() else float(number)
        return {'value': value, 'status': 'numeric'}
    # Preserve unsupported strings/numbers in both output and diagnostics.
    return {'value': cell, 'status': 'unresolved'}


def get_c_type_value(cell):
    return classify_camera_type(cell)['code']


def get_s_type_value(cell):
    return classify_speed(cell)['value']


def classify_dataframe(df):
    if df.columns.duplicated().any() or not {'camera_type', 'speed'}.issubset(df.columns):
        raise ValueError('欄位重複或缺少 camera_type / speed')
    output = df.copy(deep=True)
    codes, speeds = [], []
    report = {'status': 'ok', 'priority': [int(code) for code in c_type],
              'input_rows': len(df), 'issues': [], 'ambiguous_rows': [],
              'unclassified_rows': [], 'unresolved_speeds': []}
    methods, speed_methods = Counter(), Counter()
    for offset, (_, row) in enumerate(df.iterrows()):
        camera = classify_camera_type(row['camera_type'])
        speed = classify_speed(row['speed'])
        codes.append(camera['code'])
        speeds.append(speed['value'])
        methods[camera['status']] += 1
        speed_methods[speed['status']] += 1
        origin = {'excel_row': offset + 2,
                  'location': {col: row[col] for col in ('city', 'district', 'road_1', 'road_2', 'address', 'type')
                               if col in row.index}}
        if len(camera['matches']) > 1:
            report['ambiguous_rows'].append({**origin, 'original': row['camera_type'],
                'selected_code': camera['code'], 'matches': camera['matches']})
        if camera['status'] in ('unknown', 'missing'):
            report['unclassified_rows'].append({**origin, 'original': row['camera_type'],
                'code': 0, 'reason': camera['status']})
        if speed['status'] == 'unresolved':
            report['unresolved_speeds'].append({**origin, 'original': row['speed']})
    output['camera_type'] = pd.Series(codes, index=df.index, dtype='int64')
    output['speed'] = pd.Series(speeds, index=df.index, dtype=object)
    untouched = [col for col in df.columns if col not in ('camera_type', 'speed')]
    if len(output) != len(df) or list(output.columns) != list(df.columns) or not output[untouched].equals(df[untouched]):
        raise ValueError('分類過程變更了筆數、欄位或其他資料')
    report.update(output_rows=len(output), output_columns=list(output.columns),
                  camera_type_counts={str(code): count for code, count in sorted(Counter(codes).items())},
                  camera_methods=dict(methods), speed_methods=dict(speed_methods))
    if not len(df):
        report['issues'].append('來源沒有資料列')
    if report['unclassified_rows'] or report['unresolved_speeds'] or report['issues']:
        report['status'] = 'needs_review'
    return output, report


def data_process3(file_path=None, output_path=None):
    file_path = file_path or os.path.join(project_path, 'split_data', 'output', 'all_cities.xlsx')
    output_path = output_path or os.path.join(project_path, 'compare_data', 'input', 'gov_data.xlsx')
    with pd.ExcelFile(file_path) as workbook:
        if len(workbook.sheet_names) != 1:
            raise ValueError('分類階段預期單一合併工作表，請先確認多工作表來源')
        df = pd.read_excel(workbook, sheet_name=workbook.sheet_names[0],
                           dtype=object, keep_default_na=False)
    missing = set(REQUIRED_COLUMNS) - set(df.columns)
    if missing:
        raise ValueError(f'缺少必要欄位：{sorted(missing)}')
    output, report = classify_dataframe(df)
    report['source'] = os.fspath(file_path)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    output.to_excel(output_path, index=False)
    report_path = os.path.splitext(output_path)[0] + '.classification.json'
    with open(report_path, 'w', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2, default=str)
    if report['status'] == 'needs_review':
        warnings.warn(f'分類完成，共 {len(output)} 筆；未分類或複合速限詳見 {report_path}', stacklevel=2)
    else:
        print(f'分類完成：{output_path}，共 {len(output)} 筆')
    return report


if __name__ == "__main__":
    data_process3()
