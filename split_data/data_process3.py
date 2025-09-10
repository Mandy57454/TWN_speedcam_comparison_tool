import pandas as pd
import os


project_path = r"C:\Users\mandy.chang\PycharmProjects\TWN_speedcam_compare"

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


def get_c_type_value(cell):
    if pd.isna(cell):
        return 0  # 空值就維持原樣
    for num, keywords in c_type.items():
        for kw in keywords:
            if str(kw) in str(cell):  # 關鍵字出現在內容裡
                return int(num)
    return 0  # 沒找到關鍵字就維持原值


def get_s_type_value(cell):
    if pd.isna(cell):
        return cell  # 空值就維持原樣
    for num, keywords in s_type.items():
        for kw in keywords:
            if str(kw) in str(cell):  # 關鍵字出現在內容裡
                return int(num)
    return cell  # 沒找到關鍵字就維持原值


def data_process3():
    file_path = os.path.join(project_path, r"split_data\output\all_cities.xlsx")
    output_path = os.path.join(project_path, r"compare_data\input\gov_data.xlsx")

    df = pd.read_excel(file_path, sheet_name='Sheet1')

    # 用apply一行一行處理
    df['camera_type'] = df['camera_type'].apply(get_c_type_value)
    df['speed'] = df['speed'].apply(get_s_type_value)

    df.to_excel(output_path, index=False)


if __name__ == "__main__":
    data_process3()
