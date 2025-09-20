import pandas as pd
import os
import shutil
import pdfplumber
from openpyxl import load_workbook
from dateutil.parser import parse
from openpyxl.styles import Font
from process_taipei_data import process_taipei_data
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


# 從 PDF 中擷取所有表格
def extract_tables_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        tables = []  # 初始化表格清單
        for page_no, page in enumerate(pdf.pages):
            page_tables = page.extract_tables()
            print(f"Page {page_no + 1} extracted {len(page_tables)} tables")
            for t_no, table in enumerate(page_tables):
                print(f"  Table {t_no} on page {page_no + 1} has {len(table)} rows")
                tables.append(table)
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
    if not tables:
        return []

        # 1. 找header（同你的寫法）
    header = None
    header_idx = None
    for i, row in enumerate(tables[0]):
        if row and '編號' in row:
            header = row
            header_idx = i
            break
    if header is None:
        header = tables[0][0]
        header_idx = 0
        print("WARNING: 找不到標準表頭，改以第一行作為 header:", header)

    has_region = '區域' in header
    region_idx = header.index('區域') if has_region else None

    # 2. 扁平化
    flat_rows = []
    for tbl in tables:
        flat_rows.extend(tbl)

    # 3. 分組（依欄位數量）
    length_to_rows = {}
    for row in flat_rows:
        n = len(row)
        if n not in length_to_rows:
            length_to_rows[n] = []
        length_to_rows[n].append(row)

    # 4. 每組做資料補齊，處理 header & region
    dfs = []
    for n_cols, rows in length_to_rows.items():
        # header補足：如果比header多，補上新欄名
        this_header = list(header)
        if n_cols > len(this_header):
            # 補新欄名（用第一筆多出來的值命名）
            extra_names = []
            for i in range(len(this_header), n_cols):
                name = rows[0][i]
                if name in this_header or name in extra_names:
                    name = f"異常欄位{i + 1}"
                extra_names.append(name)
            this_header += extra_names
        elif n_cols < len(this_header):
            # header太多則截斷
            this_header = this_header[:n_cols]

        # region自動補（如原程式）
        normalized = []
        last_region = None
        for row in rows:
            cells = list(row)
            # 處理區域補值
            if has_region and region_idx is not None and len(cells) > region_idx:
                if cells[region_idx] not in (None, ''):
                    last_region = cells[region_idx]
                else:
                    cells[region_idx] = last_region

            # 若不足欄位則補None
            if len(cells) < n_cols:
                cells += [None] * (n_cols - len(cells))

            normalized.append(cells)

        df = pd.DataFrame(normalized, columns=this_header)
        df = df.apply(lambda col: col.map(try_parse_value))
        if has_region and '區域' in df.columns:
            df['區域'] = df['區域'].ffill()

        dfs.append(df)
    return dfs


def main():
    for city in cities:
        # if city == 'ChangHua':
            folder_path = os.path.join(project_path, r"selenium\city_data", city)  # 原始 PDF 資料夾
            excel_path = os.path.join(project_path, r"converter\city_data", city)  # 轉換後 Excel 資料夾
            # check download path exists
            if os.path.exists(excel_path):
                clear_folder(excel_path)  # 清空舊檔案
            else:
                os.makedirs(excel_path)  # 建立新資料夾
            # loop folder 中的所有 file
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                if filename.endswith('.pdf'):  # 若為 PDF 檔案
                    output_path = os.path.join(excel_path, os.path.splitext(filename)[0] + '.xlsx')
                    # 提取PDF中的所有表格数据
                    tables = extract_tables_from_pdf(file_path)  # 擷取 PDF 表格
                    print(f"Extracted {len(tables)} tables from {filename}.") \
                        if tables else print(f"No tables extracted from {filename}.")  # 列印結果

                    dfs = convert_to_dataframe(tables)
                    if not dfs:  # 沒有任何表
                        # 建立一個空的DataFrame，讓excel一定有一個sheet
                        empty_df = pd.DataFrame([["無資料"]], columns=["狀態"])
                        with pd.ExcelWriter(output_path) as writer:
                            empty_df.to_excel(writer, sheet_name="無資料", index=False)
                    else:
                        with pd.ExcelWriter(output_path) as writer:
                            for i, df in enumerate(dfs):
                                df.to_excel(writer, sheet_name=f"Sheet_{i + 1}", index=False)
                        # 列印完成訊息
                        print(f"PDF 中的表格已成功提取並保存到 {output_path}")

                # 原本就已經是 .xlsx or .xls 直接 copy 到 output path
                elif filename.endswith('.xlsx') or filename.endswith('.xls'):
                    output_path = os.path.join(excel_path, filename)
                    shutil.copy(file_path, output_path)
                    print(f"Excel 文件已成功複製到 {output_path}")
                # 其他不是.pdf or .xlsx or .xls 的直接 skip
                else:
                    print(f"Skipping non-PDF or non-Excel file: {filename}")


if __name__ == "__main__":
    main()
    process_taipei_data()
