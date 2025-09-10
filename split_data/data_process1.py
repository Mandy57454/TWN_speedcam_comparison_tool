import pandas as pd
import os
import shutil
import re
import xlrd
from openpyxl import Workbook

project_path = r"C:\Users\mandy.chang\PycharmProjects\TWN_speedcam_compare"

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
    '編號': {'編號', '項次', '標號', '序號'},
    '設置地點': {'設置地點', '地點', '設備設置地點', '勤務地點', '測照地點', '設置位置', '路段', '取締地點', '固定桿設置地點', '地點或路段'},
    '速限': {'速限'},
    '取締項目': {'取締項目', '設備名稱', '設備型式', '違規取締項目', '測照型式', '違規偵測項目', '器材種類', '執法項目', '取締違規項目', '功能'},
    '拍攝方向': {'拍攝方向', '測照方向', '取締行向', '方向', '測照行向'},
    '行政區': {'行政區', '地區'},
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
    # 處理欄位名稱
    df.columns = [clean_value(col) for col in df.columns]
    # 處理每個 cell
    for col in df.columns:
        df[col] = df[col].apply(clean_value)
    return df


def standardize_df(df, alias_map):
    """欄位、內容標準化"""
    df = df.rename(columns=lambda col: alias_map.get(str(col), col))  # 欄位標準化
    for col in df.columns:
        df[col] = df[col].apply(lambda v: alias_map.get(str(v), v))  # 內容標準化
    return df


def safe_sheet_name(name, existings):
    """確保Excel sheet名合法且唯一"""
    safe_name = str(name).replace('/', '_').replace('\\', '_')[:30]
    original = safe_name
    idx = 1
    while safe_name in existings:
        safe_name = f"{original}_{idx}"
        idx += 1
    return safe_name


def split_by_bianhao_sections(df):
    """根據每個'編號'區塊切割並取該行為標題"""
    positions = []
    for row_idx in range(df.shape[0]):
        for col_idx in range(df.shape[1]):
            if str(df.iat[row_idx, col_idx]) == "編號":
                positions.append(row_idx)
    if not positions:
        return {}, df  # 無"編號"則不分割

    positions.append(df.shape[0])  # 最後一段結尾
    sheets = {}
    used_names = set()

    for i in range(len(positions)-1):
        start_row = positions[i]
        end_row = positions[i+1]
        # header_row為該區塊的"編號"那一行
        header_row = df.iloc[start_row, :].tolist()
        section = df.iloc[start_row+1:end_row, :]
        if not section.empty:
            section.columns = header_row
            section = section.reset_index(drop=True)
            # sheet名可自訂，預設取 header_row 第一格
            bianhao_title = header_row[0] if header_row[0] != "編號" else f"sheet{i+1}"
            safe_name = safe_sheet_name(bianhao_title, used_names)
            sheets[safe_name] = section
            used_names.add(safe_name)
    # 原始資料僅保留第一個"編號"前的內容
    trimmed_df = df.iloc[:positions[0], :].reset_index(drop=True)
    return sheets, trimmed_df


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
    max_check = 10  # 最多往下檢查幾行，避免死循環
    check_cnt = 0
    while '設置地點' not in df.columns and len(df) > 0 and check_cnt < max_check:
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        check_cnt += 1
    return df


def reorder_columns(df):
    must_have = ['編號', '行政區', '設置地點', '速限', '取締項目', '拍攝方向', '偵測距離']
    # 1. 處理重複欄位名稱
    new_columns = []
    col_count = {}
    for col in df.columns:
        if col in must_have:
            col_count[col] = col_count.get(col, 0) + 1
            if col_count[col] == 1:
                new_columns.append(col)
            else:
                new_columns.append(f"{col}{col_count[col] - 1}")
        else:
            new_columns.append(col)
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


def convert_xls(folder_path):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.xls'):
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
        # if city == 'YiLan':
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
                        dfs = pd.read_excel(file_path, sheet_name=None)
                        all_sheets = {}

                        for sheet_name, df in dfs.items():
                            df = clean_df(df)
                            df = standardize_df(df, alias_map)
                            df = df.dropna(subset=[df.columns[0], df.columns[1]])
                            df = ensure_address_column(df)

                            # 拆分出多個分 sheet（依編號分段）
                            sheets, trimmed_df = split_by_bianhao_sections(df)
                            # 主 sheet 先加進去
                            main_sheet_name = sheet_name  # 你可以自訂這個命名方式
                            all_sheets[main_sheet_name] = reorder_columns(trimmed_df)
                            # 再把所有分段 sheet 都加進去
                            for sname, sdf in sheets.items():
                                # 可以讓 sname 加上原本 sheet 名以避免重複
                                final_name = f"{sheet_name}_{sname}" if len(dfs) > 1 else sname
                                all_sheets[final_name] = reorder_columns(sdf)

                        with pd.ExcelWriter(output_path) as writer:
                            for sname, sdf in all_sheets.items():
                                sdf.to_excel(writer, sheet_name=sname[:31], index=False)  # Excel sheet name 最長31字
                    except Exception as e:
                        print(f"讀取或寫出失敗: {filename}, 錯誤: {e}")


if __name__ == "__main__":
    data_process1()
