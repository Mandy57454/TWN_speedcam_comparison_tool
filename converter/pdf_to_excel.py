import pandas as pd
import os
import shutil
import pdfplumber
import json
import re
from collections import Counter
from statistics import median
import warnings
# Define the project path
# 動態取得專案路徑
project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Define the cities list
cities = {  # 所有城市或區域名稱
    'taipei', 'TaoYuan', 'KeeLung', 'Science_Park', 'HsinChu_web',
    'MiaoLi', 'TaiChung', 'YunLin', 'ChiaYi', 'NanTou', 'TaiNan',
    'TaiTung', 'HuaLian', 'YiLan', 'KinMen', 'Kaohsiung', 'ChiaYi_sh',
    'PengHu', 'ChangHua', 'newTaipei', 'HsinChu', 'PingTung',
}


# check folder exists and clear folder all content
def clear_folder(o_folder_path):
    # loop through all files in the folder
    for o_filename in os.listdir(o_folder_path):
        # construct the full file path
        o_file_path = os.path.join(o_folder_path, o_filename)
        try:
            if os.path.isfile(o_file_path) or os.path.islink(o_file_path):
                os.unlink(o_file_path)
            elif os.path.isdir(o_file_path):
                shutil.rmtree(o_file_path)
        except Exception as e:
            print(f'Failed to delete {o_file_path}. Reason: {e}')


HEADER_ALIASES = {
    '編號': {'編號', '設備編號', '項次', '序號', '標號', '設備'},
    '行政區': {'行政區', '區域', '地區', '鄉鎮市區'},
    '設置地點': {'設置地點', '設備設置地點', '設置位置', '地點', '取締地點',
               '路段', '地點或路段', '設置地址', '測照地點', '固定桿設置地點'},
    '取締項目': {'取締項目', '執法項目', '違規取締項目', '取締違規項目',
               '違規偵測項目', '設備名稱', '功能', '測照型式'},
    '速限': {'速限', '速限(km/h)', '速限(公里/小時)'},
    '拍攝方向': {'拍攝方向', '測照方向', '方向', '取締行向', '測照行向'},
    '座標緯度': {'座標緯度', '緯度', 'latitude', 'Latitude'},
    '座標經度': {'座標經度', '經度', 'longitude', 'Longitude'},
    '管轄單位': {'管轄單位', '管轄分局', '分局', '所屬分局'},
    '備註': {'備註', '備考'},
}
ALIAS_TO_FIELD = {alias: field for field, aliases in HEADER_ALIASES.items()
                  for alias in aliases}
CHECK_WORDS = ('超速', '闖紅燈')


def compact(value):
    if isinstance(value, float):
        if pd.isna(value):
            return ''
        if value.is_integer():
            return str(int(value))
    return re.sub(r'\s+', '', str(value)) if value is not None else ''


def header_key(value):
    text = compact(value).replace('（', '(').replace('）', ')')
    return ALIAS_TO_FIELD.get(text, text)


def find_header(rows):
    """Require several known fields; a lone '編號' is not a header."""
    for index, row in enumerate(rows):
        keys = {header_key(value) for value in row if compact(value)}
        known = keys & HEADER_ALIASES.keys()
        if len(known) >= 3 and ('設置地點' in known or '編號' in known):
            header = [compact(value) for value in row]
            end = index + 1
            # Some PDFs put only the remainder of a header on the next row.
            if end < len(rows):
                following = rows[end]
                populated = [(col, compact(value)) for col, value in enumerate(following)
                             if compact(value)]
                if populated and all(
                    col < len(header) and header[col] == '設備' and value == '編號'
                    for col, value in populated
                ):
                    for col, value in populated:
                        header[col] += value
                    end += 1
            # Excel column labels must be nonempty and unique.
            counts = Counter()
            unique = []
            for col, label in enumerate(header):
                label = label or f'未識別欄位{col + 1}'
                counts[label] += 1
                unique.append(label if counts[label] == 1 else f'{label}_{counts[label]}')
            return index, end, unique
    return None


class ExtractedTable(list):
    """List-compatible table with source geometry and extraction diagnostics."""
    def __init__(self, rows, page, bbox, strategy):
        super().__init__(rows)
        self.page = page
        self.bbox = bbox
        self.strategy = strategy


class ExtractedTables(list):
    def __init__(self):
        super().__init__()
        self.report = {'status': 'ok', 'pages': [], 'issues': []}


def remove_background_rectangles(obj):
    """White fill-only rectangles are backgrounds, not table borders."""
    if obj['object_type'] != 'rect' or obj.get('stroke'):
        return True
    color = obj.get('non_stroking_color')
    if isinstance(color, (int, float)):
        white = color >= .95
    elif isinstance(color, (tuple, list)) and len(color) in (1, 3):
        white = all(channel >= .95 for channel in color)
    elif isinstance(color, (tuple, list)) and len(color) == 4:
        white = all(channel <= .05 for channel in color)
    else:
        white = False
    return not (obj.get('fill') and white)


def fill_vertical_spans(table, rows, inherited_labels=None):
    """Fill only cells geometrically covered by a vertical merged cell.

    Never forward-fill genuine blanks or horizontally merged title cells.
    """
    header = find_header(rows)
    if not header and inherited_labels:
        header = (0, 0, inherited_labels)
    if not header:
        return rows
    start, end, labels = header
    geometry = [row.cells for row in table.rows]
    sources = [(box, rows[r][c]) for r, cells in enumerate(geometry)
               for c, box in enumerate(cells)
               if r >= end and box is not None and compact(rows[r][c])]
    result = [list(row) for row in rows]
    for r in range(end, len(rows)):
        cells = geometry[r]
        present = [box for box in cells if box is not None]
        if not present:
            continue
        top = max(box[1] for box in present)
        bottom = min(box[3] for box in present)
        if bottom <= top:
            continue
        y = (top + bottom) / 2
        for c, value in enumerate(rows[r]):
            if cells[c] is not None or compact(value) or header_key(labels[c]) == '編號':
                continue
            column_boxes = [boxes[c] for boxes in geometry[start:]
                            if c < len(boxes) and boxes[c] is not None]
            if not column_boxes:
                continue
            x0 = median(box[0] for box in column_boxes)
            x1 = median(box[2] for box in column_boxes)
            for box, text in sources:
                if (abs(box[0] - x0) < 3 and abs(box[2] - x1) < 3
                        and box[1] <= y <= box[3]):
                    result[r][c] = text
                    break
    return result


def column_signature(table, count, page_width):
    columns = []
    for col in range(count):
        boxes = [row.cells[col] for row in table.rows
                 if len(row.cells) == count and row.cells[col] is not None]
        if not boxes:
            return None
        columns.append((median(box[0] for box in boxes) / page_width,
                        median(box[2] for box in boxes) / page_width))
    return columns


def numeric_ids(rows, labels):
    keys = [header_key(label) for label in labels]
    if '編號' not in keys:
        return []
    col = keys.index('編號')
    return [int(compact(row[col])) for row in rows
            if len(row) > col and re.fullmatch(r'\d+', compact(row[col]))]


def candidate_rows(found, page_width, previous):
    result = []
    for table in found:
        rows = table.extract()
        inherited = None
        if not find_header(rows) and previous:
            labels = previous['labels']
            signature = column_signature(table, len(labels), page_width)
            ids = numeric_ids(rows, labels)
            aligned = signature and all(
                # Allow modest column resizing and page-margin shifts.
                abs(a - b) <= .02 for current, old in zip(signature, previous['columns'])
                for a, b in zip(current, old))
            # Geometry AND continuous numbering are required to inherit a header.
            if aligned and ids and ids[0] == previous['last_id'] + 1:
                inherited = labels
        filled = fill_vertical_spans(table, rows, inherited)
        result.append([inherited] + filled if inherited else filled)
    return result


def candidate_quality(tables, source_text):
    """Prefer intact records over fragmented rows, without rewarding row count."""
    recognized = records = fragments = populated = 0
    for rows in tables:
        header = find_header(rows)
        if not header:
            continue
        _, end, labels = header
        keys = [header_key(label) for label in labels]
        recognized += len(set(keys) & HEADER_ALIASES.keys())
        id_col = keys.index('編號') if '編號' in keys else None
        for row in rows[end:]:
            if not any(compact(value) for value in row):
                continue
            if id_col is not None:
                if re.fullmatch(r'\d+', compact(row[id_col])):
                    records += 1
                    populated += sum(bool(compact(value)) for value in row)
                else:
                    fragments += 1
            else:
                records += 1
                populated += sum(bool(compact(value)) for value in row)
    extracted = compact(''.join(str(value or '') for rows in tables
                                for row in rows for value in row))
    missing_words = sum(max(0, source_text.count(word) - extracted.count(word))
                        for word in CHECK_WORDS)
    return (bool(recognized), -missing_words, records, -fragments, populated, recognized)


def table_source_text(page, found):
    """Compare data-area text, excluding page titles and column headings."""
    parts = []
    for table in found:
        bbox = list(table.bbox)
        header = find_header(table.extract())
        if header and header[1] < len(table.rows):
            boxes = [box for box in table.rows[header[1]].cells if box is not None]
            if boxes:
                bbox[1] = max(box[1] for box in boxes)
        if bbox[3] > bbox[1]:
            parts.append(compact(page.crop(tuple(bbox)).extract_text() or ''))
    return ''.join(parts)


def extract_tables_from_pdf(pdf_path):
    tables = ExtractedTables()
    previous = None
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            source_content = page.extract_text() or ''
            source_text = compact(source_content)
            page_report = {'page': page_no, 'issues': [], 'candidates': []}
            candidates = []
            strategies = [('lines', page, {})]
            if any(not remove_background_rectangles(rect) for rect in page.rects):
                strategies.append(('without_backgrounds',
                                   page.filter(remove_background_rectangles), {}))
            for name, surface, settings in strategies:
                try:
                    found = surface.find_tables(table_settings=settings)
                    rows = candidate_rows(found, page.width, previous)
                    score = candidate_quality(rows, table_source_text(page, found))
                    candidates.append((score, name, found, rows))
                    page_report['candidates'].append({'strategy': name, 'quality': score})
                except Exception as error:
                    page_report['issues'].append(f'{name}: {error}')
            # Text alignment is a fallback; do not let it replace a healthy grid.
            best = max(candidates, key=lambda item: item[0]) if candidates else None
            if best is None or not best[0][0] or best[0][1] < 0 or best[0][3] < 0:
                try:
                    found = page.find_tables(table_settings={
                        'vertical_strategy': 'text', 'horizontal_strategy': 'text'})
                    rows = candidate_rows(found, page.width, previous)
                    score = candidate_quality(rows, table_source_text(page, found))
                    candidates.append((score, 'text', found, rows))
                    page_report['candidates'].append({'strategy': 'text', 'quality': score})
                except Exception as error:
                    page_report['issues'].append(f'text: {error}')
            if candidates:
                score, name, found, rows = max(candidates, key=lambda item: item[0])
                page_report.update(strategy=name, table_count=len(rows))
                for table, values in zip(found, rows):
                    tables.append(ExtractedTable(values, page_no, table.bbox, name))
                # Inheritance is restricted to the immediately preceding page.
                previous = None
                if found:
                    header = find_header(rows[-1])
                    if header:
                        _, end, labels = header
                        ids = numeric_ids(rows[-1][end:], labels)
                        columns = column_signature(found[-1], len(labels), page.width)
                        if ids and columns:
                            previous = {'labels': labels, 'columns': columns,
                                        'last_id': ids[-1]}
                if not score[0]:
                    page_report['issues'].append('無法辨識可靠表頭，保留原始表格待檢查')
                if score[1] < 0:
                    page_report['issues'].append('擷取結果缺少來源中的執法關鍵文字')
                if score[3] < 0:
                    page_report['issues'].append('存在無設備編號的資料列，未自動拼接')
            else:
                previous = None
                page_report['issues'].append('無法擷取表格')
            if not source_text:
                page_report['issues'].append('頁面無可用文字，可能需要 OCR')
            if not any(item.page == page_no for item in tables):
                page_report['issues'].append('頁面沒有可輸出的表格')
                if source_content:
                    tables.append(ExtractedTable([[line] for line in source_content.splitlines()],
                                                 page_no, page.bbox, 'raw_text'))
            tables.report['pages'].append(page_report)
            print(f"Page {page_no}: {page_report.get('strategy', 'none')}, "
                  f"{page_report.get('table_count', 0)} tables")
            for issue in page_report['issues']:
                warnings.warn(f'第 {page_no} 頁：{issue}', stacklevel=2)
    if any(page['issues'] for page in tables.report['pages']):
        tables.report['status'] = 'needs_review'
    return tables


# 嘗試將字串轉為整數、浮點數
def try_parse_value(val):
    # 如果已经不是字符串，直接返回原值
    if not isinstance(val, str):  # 若已經不是字串，直接回傳原值
        return val
    text = val.strip()  # 去除前後空白
    if text.isdigit():  # 若全為數字字元
        return int(text)  # 回傳整數
    try:
        f = float(text)  # 嘗試轉為浮點數
        return f
    except (ValueError, OverflowError):
        return val  # 失敗則回傳原字串


# 將擷取的原始表格資料轉換為 pandas DataFrame
def convert_to_dataframe(tables):
    groups = {}
    for table_no, table in enumerate(tables):
        header = find_header(table)
        if not header:
            # Unknown layouts stay isolated; never borrow another table's header.
            width = max((len(row) for row in table), default=0)
            if width:
                groups[('raw', table_no)] = {
                    'labels': [f'未識別欄位{i + 1}' for i in range(width)],
                    'rows': list(table), 'sources': [table_no], 'recognized': False}
            continue
        start, end, labels = header
        key = tuple(header_key(label) for label in labels)
        group = groups.setdefault(key, {'labels': labels, 'rows': [],
                                        'sources': [], 'recognized': True})
        group['sources'].append(table_no)
        index = end
        while index < len(table):
            repeated = find_header(table[index:index + 2])
            if repeated and repeated[0] == 0:
                _, skip, new_labels = repeated
                if tuple(header_key(label) for label in new_labels) == key:
                    index += skip
                    continue
                # A different schema starts a new section, even on the same page.
                tail = ExtractedTable(table[index:], getattr(table, 'page', None),
                                      getattr(table, 'bbox', None),
                                      getattr(table, 'strategy', 'unknown'))
                for frame in convert_to_dataframe([tail]):
                    isolated_key = ('section', table_no, index, len(groups))
                    groups[isolated_key] = {'labels': list(frame.columns),
                        'rows': frame.values.tolist(), 'sources': [table_no],
                        'recognized': frame.attrs.get('recognized', False)}
                break
            row = table[index]
            if any(compact(value) for value in row):
                group['rows'].append(list(row))
            index += 1
    frames = []
    for group in groups.values():
        width = max([len(group['labels'])] + [len(row) for row in group['rows']])
        labels = group['labels'] + [f'未識別欄位{i + 1}'
                                   for i in range(len(group['labels']), width)]
        frame = pd.DataFrame([row + [None] * (width - len(row))
                              for row in group['rows']], columns=labels)
        frame = frame.apply(lambda col: col.map(try_parse_value))
        frame.attrs.update(recognized=group['recognized'], source_tables=group['sources'])
        frames.append(frame)
    return frames


def conversion_report(tables, frames):
    report = getattr(tables, 'report', {'status': 'ok', 'pages': [], 'issues': []})
    report['tables'] = [{'index': index, 'page': getattr(table, 'page', None),
                         'bbox': getattr(table, 'bbox', None),
                         'strategy': getattr(table, 'strategy', 'unknown')}
                        for index, table in enumerate(tables)]
    report['sheets'] = []
    source_ids = Counter()
    output_ids = Counter()
    for table in tables:
        header = find_header(table)
        if header:
            source_ids.update(numeric_ids(table[header[1]:], header[2]))
    for index, frame in enumerate(frames, 1):
        keys = [header_key(label) for label in frame.columns]
        issues = []
        if not frame.attrs.get('recognized'):
            issues.append('未辨識格式，保留原始資料')
        if frame.empty:
            issues.append('沒有資料列')
        if '編號' in keys:
            output_ids.update(numeric_ids(frame.values.tolist(), list(frame.columns)))
            ids = frame.iloc[:, keys.index('編號')].map(compact)
            numeric = ids.str.fullmatch(r'\d+')
            if not numeric.all():
                issues.append(f'{int((~numeric).sum())} 列缺少數字設備編號，保留待檢查')
            if ids[numeric].duplicated().any():
                issues.append('設備編號重複，需確認是否跨區重新編號')
        for required in ('設置地點', '行政區'):
            if required in keys:
                missing = frame.iloc[:, keys.index(required)].map(
                    lambda value: pd.isna(value) or not compact(value)).sum()
                if missing:
                    issues.append(f'{required}有 {int(missing)} 列空白')
        for coordinate, bound in (('座標緯度', 90), ('座標經度', 180)):
            if coordinate in keys:
                values = pd.to_numeric(frame.iloc[:, keys.index(coordinate)], errors='coerce')
                invalid = (values.isna() | (values.abs() > bound)).sum()
                if invalid:
                    issues.append(f'{coordinate}有 {int(invalid)} 列缺漏或不合法')
        report['sheets'].append({'sheet': f'Sheet_{index}', 'rows': len(frame),
                                'source_tables': frame.attrs.get('source_tables', []),
                                'issues': issues})
        if issues:
            report['status'] = 'needs_review'
    report['device_count'] = {'extracted': sum(source_ids.values()),
                              'output': sum(output_ids.values())}
    if source_ids != output_ids:
        report['issues'].append('轉換前後的設備編號或出現次數不一致')
        report['status'] = 'needs_review'
    # Removing titles/headers must never remove enforcement text silently.
    before = compact(''.join(str(value or '') for table in tables
                             for row in table for value in row))
    after = compact(''.join(str(value) for frame in frames
                            for row in frame.fillna('').values for value in row))
    for word in CHECK_WORDS:
        if before.count(word) != after.count(word):
            report['issues'].append(f'{word}文字數改變：{before.count(word)} → {after.count(word)}')
            report['status'] = 'needs_review'
    return report


def convert_pdf(pdf_path, output_path):
    """Convert one PDF without clearing any directories; return diagnostics."""
    tables = extract_tables_from_pdf(pdf_path)
    frames = convert_to_dataframe(tables)
    report = conversion_report(tables, frames)
    report['source'] = os.fspath(pdf_path)
    if not frames:
        frames = [pd.DataFrame([['無資料']], columns=['狀態'])]
        report['status'] = 'needs_review'
        report['issues'].append('沒有可解析的表格，請檢查來源或使用 OCR')
    with pd.ExcelWriter(output_path) as writer:
        for index, frame in enumerate(frames, 1):
            sheet_name = '無資料' if not tables else f'Sheet_{index}'
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    report_path = os.path.splitext(output_path)[0] + '.conversion.json'
    with open(report_path, 'w', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
    if report['status'] == 'ok':
        print(f'PDF 表格已保存到 {output_path}')
    else:
        warnings.warn(f'轉換結果需要檢查：{output_path}；詳見 {report_path}', stacklevel=2)
    return report


def main():
    for city in cities:
        try:
            folder_path = os.path.join(project_path, r"selenium\city_data", city)  # 原始 PDF 資料夾
            excel_path = os.path.join(project_path, r"converter\city_data", city)  # 轉換後 Excel 資料夾
            
            # 檢查原始資料夾是否存在
            if not os.path.exists(folder_path):
                print(f"警告: 找不到資料夾 {folder_path}，跳過城市 {city}")
                continue
            
            # check download path exists
            if os.path.exists(excel_path):
                clear_folder(excel_path)  # 清空舊檔案
            else:
                os.makedirs(excel_path)  # 建立新資料夾
            
            print(f"正在處理城市: {city}")
            # loop folder 中的所有 file
            for filename in os.listdir(folder_path):
                try:
                    file_path = os.path.join(folder_path, filename)
                    if filename.endswith('.pdf'):  # 若為 PDF 檔案
                        output_path = os.path.join(excel_path, os.path.splitext(filename)[0] + '.xlsx')
                        convert_pdf(file_path, output_path)

                    # 原本就已經是 .xlsx or .xls 直接 copy 到 output path
                    elif filename.endswith('.xlsx') or filename.endswith('.xls'):
                        output_path = os.path.join(excel_path, filename)
                        shutil.copy(file_path, output_path)
                        print(f"Excel 文件已成功複製到 {output_path}")
                    # 其他不是.pdf or .xlsx or .xls 的直接 skip
                    else:
                        print(f"Skipping non-PDF or non-Excel file: {filename}")

                except Exception as e:
                    print(f"處理檔案 {filename} 時發生錯誤: {e}")
                    continue
            
            print(f"城市 {city} 處理完成")
            
        except Exception as e:
            print(f"處理城市 {city} 時發生錯誤: {e}")
            continue


if __name__ == "__main__":
    main()
    # process_taipei_data()
