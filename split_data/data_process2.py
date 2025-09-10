import pandas as pd
import os


project_path = r"C:\Users\mandy.chang\PycharmProjects\TWN_speedcam_compare"

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


def ensure_district(x, remark, z, city):
    district = ''
    if pd.isna(x) or x == '':
        if (pd.isna(remark) or remark == '') and (pd.isna(z) or z == ''):
            return district
        z_str = str(z)
        remark_str = str(remark)
        if '區' in remark_str:
            district = remark_str
        elif '臺北市' in city:
            district = z_str
        else:
            return district
    else:
        x = str(x)
        district = x
    if '分局' in district:
        district = ''
    if any(key in district for key in ["一", "二"]):
        district = district[:-1]
    for sep in ["、", "區"]:
        if sep in district:
            district = district.split(sep)[0]
    if (not any(key in district for key in ["鄉", "鎮", "市", "區"]) or
            any(key in district for key in ["前鎮", "左鎮", "新市", "平鎮"])):
        district = district + "區"
        # print(district)
    return district


def check_address(city, district, address):
    address_str = str(address)
    if not pd.isna(city):
        city_str = str(city)
        if city_str in address_str:
            address_str = address_str.replace(city_str, "", 1)
    if not pd.isna(district) or district == '':
        district_str = str(district)
        if district_str in address_str:
            address_str = address_str.replace(district_str, "", 1)
    # print(address_str)
    return address_str


def safe_int(x):
    if x in [None, '', 'nan'] or (isinstance(x, float) and pd.isna(x)):
        return ''
    try:
        return int(x)
    except (ValueError, TypeError):
        return x


def check_camera_type(filename, sheet_name):
    if 'mobile' in filename or '移動式' in sheet_name:
        camera_type = 5
    else:
        camera_type = 0
    return camera_type


def check_banned(ban, remark, distance):
    if not pd.isna(distance) or distance == '':
        return 964
    if pd.isna(ban) or ban == '':
        return remark
    else:
        return ban


def combine_excel():
    folder_path = os.path.join(project_path, r'split_data\input\all_cities')  # 需要combine的path
    data_combined = os.path.join(project_path, r'split_data\input\all_cities.xlsx')  # combine後的path

    all_dfs = []
    for filename in os.listdir(folder_path):
        # if filename.endswith('.xlsx') or filename.endswith('.xls'):
            if filename == 'mio_data.xlsx':
                continue
            file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(file_path)
            df['Source_File'] = filename
            all_dfs.append(df)

    combined_df = pd.concat(all_dfs, ignore_index=True)
    combined_df.to_excel(data_combined, index=False)
    print(f"All files have been successfully merged into {data_combined}")


def data_process2():
    empty_city_list = []  # 記錄 new_rows 為空的城市
    for folder_name, city_name in cities.items():
        # if folder_name == 'Kaohsiung':
            folder_path = os.path.join(project_path, r"split_data\standard_in", folder_name)
            output_folder = os.path.join(project_path, r"split_data\input\all_cities")
            columns = ['city', 'district', 'road_1', 'road_2', 'note', 'address', 'camera_type', 'speed', 'type']
            ex_df = pd.DataFrame(columns=columns)
            new_rows = []
            for filename in os.listdir(folder_path):
                if filename.endswith('.xlsx'):
                    # camera_type = check_camera_type(filename)
                    file_path = os.path.join(folder_path, filename)
                    dfs = pd.read_excel(file_path, sheet_name=None)
                    for sheet_name, df in dfs.items():
                        camera_type = check_camera_type(filename, sheet_name)
                        for index, row in df.iterrows():
                            row_dict = row.to_dict()
                            district = ensure_district(row_dict.get('行政區', ''), row_dict.get('備註', ''),
                                                       row_dict.get('轄區', ''), city_name)
                            new_row = {
                                'city': city_name,
                                'district': district,
                                'road_1': '',
                                'road_2': '',
                                'note': '',
                                'address': check_address(city_name, district, row_dict.get('設置地點', '')),
                                'camera_type': (check_banned(row_dict.get('取締項目', ''), row_dict.get('備註', ''),
                                                             row_dict.get('偵測距離', ''))
                                                if camera_type == 0 else camera_type),
                                'speed': safe_int(row_dict.get('速限', '')),
                                'type': ''
                            }
                            # print(new_row)
                            new_rows.append(new_row)

            # 判斷 new_rows 為空
            if not new_rows:
                print(f"[!] {folder_name}（{city_name}）的 new_rows 為空，沒有輸出！")
                empty_city_list.append((folder_name, city_name))
            else:
                df_new = pd.DataFrame(new_rows)
                # 新增這一行印 city
                print(f"正在 concat {folder_name}（{city_name}）共 {len(df_new)} 筆資料")
                ex_df = pd.concat([ex_df, df_new], ignore_index=True)
                ex_df.to_excel(os.path.join(output_folder, f"{folder_name}.xlsx"), index=False)
    # 執行結束後報告
    if empty_city_list:
        print("=== 以下城市 new_rows 為空，沒有產生 Excel ===")
        for folder_name, city_name in empty_city_list:
            print(f"{folder_name}（{city_name}）")
    else:
        print("全部城市都有資料！")
    combine_excel()


if __name__ == "__main__":
    data_process2()

