import pandas as pd
import os
import shutil
import re
import json
import warnings
from collections import Counter
from openpyxl import Workbook

# 動態取得專案路徑
project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

cities = {
    'taipei', 'TaoYuan', 'KeeLung', 'Science_Park', 'HsinChu_web',
    'MiaoLi', 'TaiChung', 'YunLin', 'ChiaYi', 'NanTou', 'TaiNan',
    'TaiTung', 'HuaLian', 'YiLan', 'KinMen', 'Kaohsiung', 'ChiaYi_sh',
    'PengHu', 'ChangHua', 'newTaipei', 'HsinChu', 'PingTung',
}

camera_type = {
    'tech': {'科技執法'},
    'mobile': {'移動式', '非固定式'},
    'fixed': {'固定式', '固定桿'},
    'average': {'區間'},
}

fields = {
    '編號': {'編號', '項次', '標號', '序號', '編', '設備編號', '設備'},
    '設置地點': {'設置地點', '地點', '設備設置地點', '勤務地點', '測照地點', '設置位置', '路段', '取締地點', '固定桿設置地點', '地點或路段', '設置地址'},
    '速限': {'速限'},
    '取締項目': {'取締項目', '設備名稱', '設備型式', '違規取締項目', '測照型式', '違規偵測項目', '器材種類', '執法項目', '取締違規項目', '功能'},
    '拍攝方向': {'拍攝方向', '測照方向', '取締行向', '方向', '測照行向'},
    '行政區': {'行政區', '地區', '區域', '鄉鎮市區'},
    '偵測距離': {'偵測距離', '偵測長度', '偵側長度'},
}

# 建立欄位別名對應表
alias_map = {}
for key, alias_set in fields.items():
    for alias in alias_set:
        alias_map[alias] = key


def clean_value(val):
    if pd.isnull(val):
        return val
    return re.sub(r'[\s\u3000]+', '', str(val))  # 去掉所有空白、全形空白、換行、tab


def clean_df(df):
    df = df.copy()
    df.columns = [clean_value(col) for col in df.columns]
    # Work by position so duplicate column labels do not lose data.
    return df.apply(lambda col: col.map(clean_value))


def standardize_df(df, alias_map):
    """Standardize labels only; data values are never column aliases."""
    return df.rename(columns=lambda col: alias_map.get(str(clean_value(col)), clean_value(col)))


def safe_sheet_name(name, existings):
    """確保Excel sheet名合法且唯一"""
    original = re.sub(r'[\\/*?:\[\]]', '_', str(name)).strip("'") or 'Sheet'
    safe_name = original[:31]
    idx = 1
    while safe_name.casefold() in {str(value).casefold() for value in existings}:
        suffix = f'_{idx}'
        safe_name = original[:31 - len(suffix)] + suffix
        idx += 1
    return safe_name


def split_by_bianhao_sections(df):
    """Compatibility helper: split actual headers, not cells containing 編號."""
    sections, _ = parse_sheet(df)
    return ({f'sheet{index}': section['frame']
             for index, section in enumerate(sections[1:], 1)}, sections[0]['frame'])


def clear_folder(o_folder_path):
    """刪除資料夾下所有檔案/資料夾"""
    for o_filename in os.listdir(o_folder_path):
        o_file_path = os.path.join(o_folder_path, o_filename)
        try:
            if os.path.isfile(o_file_path) or os.path.islink(o_file_path):
                os.unlink(o_file_path)
            elif os.path.isdir(o_file_path):
                shutil.rmtree(o_file_path)
        except Exception as e:
            print(f'Failed to delete {o_file_path}. Reason: {e}')


def get_camera_type_from_filename(filename, camera_type_dict):
    for key, keywords in camera_type_dict.items():
        for keyword in keywords:
            if keyword in filename:
                return key
    return None


def get_unique_filename(folder, base_name):
    """給予不重複的檔名"""
    name, ext = os.path.splitext(base_name)
    candidate = base_name
    i = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{name}{i}{ext}"
        i += 1
    return candidate


def ensure_address_column(df):
    """Only promote a confirmed header; leave unknown/mixed layouts intact."""
    if '設置地點' in standardize_df(df, alias_map).columns:
        return standardize_df(df, alias_map)
    sections, _ = parse_sheet(df)
    if len(sections) == 1 and sections[0]['recognized']:
        return sections[0]['frame']
    return df


def reorder_columns(df):
    df = df.copy()
    must_have = ['編號', '行政區', '設置地點', '速限', '取締項目', '拍攝方向', '偵測距離']
    # 1. 處理重複欄位名稱
    new_columns = []
    used = set()
    reserved = {str(col) for col in df.columns}
    for index, col in enumerate(df.columns):
        base = str(col) if pd.notna(col) and str(col) else f'原始欄位{index + 1}'
        candidate = base
        suffix = 1
        if candidate in used:
            candidate = f'{base}{suffix}'
            suffix += 1
            # Generated labels must also avoid later original labels.
            while candidate in reserved or candidate in used:
                candidate = f'{base}{suffix}'
                suffix += 1
        new_columns.append(candidate)
        used.add(candidate)
    df.columns = new_columns
    # 2. 確保 must_have 欄位一定都在，如果缺少則加空欄
    for col in must_have:
        if col not in df.columns:
            df[col] = ''
    # 3. 其他原本就有的欄位
    other_cols = [col for col in df.columns if col not in must_have]
    new_order = must_have + other_cols
    # 4. 依照新順序重新排列
    return df[new_order]


def is_empty(value):
    return pd.isna(value) or str(value).strip() == ''


def header_labels(row):
    labels = [alias_map.get(str(value), value) for value in row]
    known = {value for value in labels if isinstance(value, str)} & fields.keys()
    if len(known) >= 2 and ('設置地點' in known or '編號' in known):
        return [value if not is_empty(value) else f'原始欄位{index + 1}'
                for index, value in enumerate(labels)]
    return None


def parse_sheet(df):
    """Identify actual headers and preserve every nonblank data/unknown row.

    Preambles are retained in diagnostics. Unknown layouts and continuation
    rows are kept intact rather than guessed, forward-filled or discarded.
    """
    has_labels = any(isinstance(col, str) and not col.startswith('Unnamed:')
                     for col in df.columns)
    cleaned = clean_df(df)
    rows = cleaned.values.tolist()
    if has_labels:
        rows.insert(0, list(cleaned.columns))
    sections, preamble = [], []
    labels, body, locations = None, [], []

    def finish():
        if labels is not None:
            sections.append({'frame': pd.DataFrame(body, columns=labels),
                             'recognized': True, 'source_rows': list(locations)})

    index = 0
    while index < len(rows):
        row = rows[index]
        if all(is_empty(value) for value in row):
            index += 1
            continue
        candidate = header_labels(row)
        if candidate is not None:
            if labels != candidate:
                finish()
                labels, body, locations = candidate, [], []
            # '設備' above '編號' is a two-line header, not an orphan record.
            if index + 1 < len(rows):
                following = [(col, value) for col, value in enumerate(rows[index + 1])
                             if not is_empty(value)]
                if following and all(str(row[col]) == '設備' and str(value) == '編號'
                                     for col, value in following):
                    index += 1
        elif labels is None:
            preamble.append({'row': index + 1,
                             'values': [None if is_empty(value) else value for value in row]})
        else:
            body.append(row)
            locations.append(index + 1)
        index += 1
    finish()
    if not sections:
        nonblank = [(index + 1, row) for index, row in enumerate(rows)
                    if any(not is_empty(value) for value in row)]
        sections = [{'frame': pd.DataFrame([row for _, row in nonblank],
                        columns=[f'原始欄位{index + 1}' for index in range(df.shape[1])]),
                     'recognized': False, 'source_rows': [index for index, _ in nonblank]}]
        preamble = []  # The unknown rows already remain in the output sheet.
    return sections, preamble


def data_snapshot(frames):
    ids, keywords = Counter(), Counter({'超速': 0, '闖紅燈': 0})
    rows = 0
    for frame in frames:
        rows += len(frame)
        if '編號' in frame.columns:
            column = frame.iloc[:, list(frame.columns).index('編號')]
            for value in column:
                if not is_empty(value):
                    token = str(clean_value(value))
                    ids[re.sub(r'^(\d+)\.0+$', r'\1', token)] += 1
        for row in frame.itertuples(index=False, name=None):
            for value in row:
                if not is_empty(value):
                    text = str(clean_value(value))
                    for word in keywords:
                        keywords[word] += text.count(word)
    return {'rows': rows, 'ids': ids, 'keywords': keywords}


def section_issues(section):
    frame = section['frame']
    issues = []
    if not section['recognized']:
        issues.append('無法辨識表頭，原始資料保留待檢查')
        return issues
    if frame.empty:
        issues.append('沒有資料列')
    if frame.columns.duplicated().any():
        issues.append('欄位名稱重複，已保留所有欄位並以尾碼區分')
    for name in ('設置地點', '編號', '行政區'):
        if name not in frame.columns:
            if name == '設置地點':
                issues.append('缺少設置地點欄位')
            continue
        values = frame.iloc[:, list(frame.columns).index(name)]
        missing = [section['source_rows'][index] for index, value in enumerate(values)
                   if is_empty(value)]
        if missing:
            issues.append(f'{name}空白，原始列號：{missing}')
    return issues


def standardize_sheets(dfs, converter_report=None):
    """Pure workbook transformation; diagnostics never silently discard rows."""
    output, baseline = {}, []
    report = {'status': 'ok', 'issues': [], 'sheets': [],
              'converter': converter_report or {'status': 'not_provided'}}
    if converter_report is not None and converter_report.get('status') != 'ok':
        report['issues'].append('converter 結果待檢查，已保留來源診斷資訊')
    used = set()
    for sheet_name, df in dfs.items():
        try:
            sections, preamble = parse_sheet(df)
        except Exception as error:
            # Keep the entire original sheet when parsing fails.
            raw = clean_df(df)
            sections = [{'frame': raw, 'recognized': False,
                         'source_rows': list(range(1, len(raw) + 1))}]
            preamble = []
            report['issues'].append(f'{sheet_name} 解析失敗，保留原表：{error}')
        for index, section in enumerate(sections):
            name = sheet_name if index == 0 else (
                f'{sheet_name}_sheet{index}' if len(dfs) > 1 else f'sheet{index}')
            name = safe_sheet_name(name, used)
            used.add(name)
            baseline.append(section['frame'])
            output[name] = reorder_columns(section['frame'])
            issues = section_issues(section)
            if index == 0 and any(sum(value is not None for value in row['values']) > 1
                                  for row in preamble):
                issues.append('表頭之前有多欄內容，已保留於前置資料診斷中')
            report['sheets'].append({'sheet': name, 'source_sheet': sheet_name,
                'recognized': section['recognized'], 'rows': len(section['frame']),
                'source_rows': section['source_rows'], 'issues': issues,
                'preamble': preamble if index == 0 else []})
    before, after = data_snapshot(baseline), data_snapshot(output.values())
    report['counts'] = {
        'input_rows': before['rows'], 'output_rows': after['rows'],
        'input_ids': sum(before['ids'].values()), 'output_ids': sum(after['ids'].values()),
        'input_keywords': dict(before['keywords']), 'output_keywords': dict(after['keywords'])}
    if before != after:
        report['issues'].append('處理前後的資料列、設備編號或執法關鍵文字不一致')
        report['missing_ids'] = list((before['ids'] - after['ids']).elements())
        report['additional_ids'] = list((after['ids'] - before['ids']).elements())
    if converter_report and isinstance(converter_report.get('device_count'), dict):
        expected = converter_report['device_count'].get('output')
        if expected is not None and expected != sum(before['ids'].values()):
            report['issues'].append('輸入設備數與 converter 診斷檔不一致，請確認檔案版本')
    if report['issues'] or any(sheet['issues'] for sheet in report['sheets']):
        report['status'] = 'needs_review'
    return output, report


def process_workbook(file_path, output_path):
    """Process one workbook without clearing or changing other files."""
    diagnostic_path = os.path.splitext(file_path)[0] + '.conversion.json'
    converter_report = None
    if os.path.exists(diagnostic_path):
        try:
            with open(diagnostic_path, encoding='utf-8') as stream:
                converter_report = json.load(stream)
            if not isinstance(converter_report, dict):
                raise ValueError('診斷檔必須是 JSON object')
        except (OSError, ValueError) as error:
            converter_report = {'status': 'needs_review', 'issues': [f'診斷檔讀取失敗：{error}']}
    # header=None preserves the first row even for unknown or title-first layouts.
    with pd.ExcelFile(file_path) as workbook:
        dfs = pd.read_excel(workbook, sheet_name=None, header=None, dtype=object,
                            keep_default_na=False)
    sheets, report = standardize_sheets(dfs, converter_report)
    report['source'] = os.fspath(file_path)
    if not sheets:
        raise ValueError('來源活頁簿沒有工作表')
    with pd.ExcelWriter(output_path) as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
    report_path = os.path.splitext(output_path)[0] + '.processing.json'
    with open(report_path, 'w', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2, default=str)
    if report['status'] == 'ok':
        print(f'標準化完成：{output_path}')
    else:
        warnings.warn(f'標準化結果待檢查：{output_path}；詳見 {report_path}', stacklevel=2)
    return report


def convert_xls(folder_path):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.xls'):
            import xlrd  # Needed only for legacy .xls files.
            file_path = os.path.join(folder_path, filename)
            wb_xls = xlrd.open_workbook(file_path)
            for sheet_idx in range(wb_xls.nsheets):
                sheet = wb_xls.sheet_by_index(sheet_idx)
                wb_xlsx = Workbook()
                ws = wb_xlsx.active
                ws.title = sheet.name

                for row in range(sheet.nrows):
                    ws.append(sheet.row_values(row))

                # 輸出檔名同原名，副檔名改為.xlsx
                out_file = os.path.splitext(filename)[0] + '.xlsx'
                # 若有多個工作表，檔名後加 sheet name
                if wb_xls.nsheets > 1:
                    out_file = os.path.splitext(filename)[0] + f'_{sheet.name}.xlsx'
                out_path = os.path.join(folder_path, out_file)
                wb_xlsx.save(out_path)
                print(f"轉換完成：{out_path}")


def data_process1():
    for city in cities:
        # if city == 'TaiChung':
            folder_path = os.path.join(project_path, r"converter\city_data", city)
            output_folder = os.path.join(project_path, r"split_data\standard_in", city)
            # 清空 output_folder
            if os.path.exists(output_folder):
                clear_folder(output_folder)
            else:
                os.makedirs(output_folder)
            os.makedirs(output_folder, exist_ok=True)
            convert_xls(folder_path)
            # 讀取所有檔案
            # 以相同的檔名命名各種相關檔案
            for filename in os.listdir(folder_path):
                if filename.endswith('.xlsx'):
                    file_path = os.path.join(folder_path, filename)
                    cam_type = get_camera_type_from_filename(filename, camera_type)
                    if cam_type:
                        output_filename = get_unique_filename(output_folder, f"{cam_type}.xlsx")
                    else:
                        output_filename = get_unique_filename(output_folder, filename)
                    output_path = os.path.join(output_folder, output_filename)
                    try:
                        process_workbook(file_path, output_path)
                    except Exception as e:
                        print(f"讀取或寫出失敗: {filename}, 錯誤: {e}")


if __name__ == "__main__":
    data_process1()
