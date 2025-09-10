import os
import pandas as pd
import shutil
from step_split import (step1_split, step2_split, step3_split, step4_split, step5_split, step6_split, step7_split,
                        step8_split, step9_split, step10_split, step11_split, step12_split, step13_split, step14_split,
                        step15_split, step16_split, step17_split, step18_split)
from mio_data_process import mio_data_process
from data_process1 import data_process1
from data_process2 import data_process2
from data_process3 import data_process3
import unicodedata

project_path = r"C:\Users\mandy.chang\PycharmProjects\TWN_speedcam_compare"


def clear_folder(o_folder_path):
    for o_filename in os.listdir(o_folder_path):
        o_file_path = os.path.join(o_folder_path, o_filename)
        try:
            if os.path.isfile(o_file_path) or os.path.islink(o_file_path):
                os.unlink(o_file_path)
            elif os.path.isdir(o_file_path):
                shutil.rmtree(o_file_path)
        except Exception as e:
            print(f'Failed to delete {o_file_path}. Reason: {e}')


# Normalize Unicode strings to NFC form
def normalize_to_basic_characters(text):
    return unicodedata.normalize('NFKC', text)


def remove_tabs_newlines_spaces(text):
    """
    Remove all tab and newline characters from the given text.

    Args:
        text (str): The input string from which to remove tab and newline characters.

    Returns:
        str: The resulting string with all tab and newline characters removed.
    """
    # 使用 replace 方法逐个清除
    text = text.replace('\t', '')
    text = text.replace('\n', '')
    text = text.replace(' ', '')
    return text


def remove_after_comma(text):
    # 先處理逗号
    result = text.split(",")[0]
    # 再處理连字符
    result = result.split("-")[0]
    # 最後處理
    result = result.split("【")[0]
    return result


def main(input_path, output_path, invalid_path):

    split_steps = [step1_split, step2_split, step3_split, step4_split, step5_split, step6_split, step7_split,
                   step8_split, step9_split, step10_split, step11_split, step12_split, step13_split, step14_split,
                   step15_split, step16_split, step17_split, step18_split]

    for filename in os.listdir(input_path):
        if filename.endswith('.xlsx'):
            seen_addresses = set()

            file_path = os.path.join(input_path, filename)
            df = pd.read_excel(file_path)
            df_valid = pd.DataFrame(columns=df.columns)
            df_invalid = pd.DataFrame(columns=df.columns)

            items = {'city', 'district', 'road_1', 'road_2', 'note'}
            for item in items:
                df_valid[item] = df_valid[item].astype(str)

            for i, row in df.iterrows():
                parts = [str(row['city']), str(row['district']), str(row['road_1']), str(row['address'])]
                combined_address = ''.join([part for part in parts if part and pd.notna(
                    part) and part != 'nan' and part != 'Z' and part != '第三方' and part.strip()])
                # combined_address = '臺東縣臺東市中興路與利嘉路口'
                combined_address_normalized = (  # 將地址的贅詞刪除
                    remove_after_comma(remove_tabs_newlines_spaces(normalize_to_basic_characters(combined_address))))
                print(combined_address_normalized)
                if combined_address_normalized in seen_addresses:
                    continue  # 如果地址已存在，跳过该行处理

                seen_addresses.add(combined_address_normalized)  # 将地址加入已处理集合

                split_result = None
                for step in split_steps:
                    split_result = step(combined_address_normalized)
                    if split_result:
                        break

                if split_result:
                    # print(split_result)
                    # print(repr(split_result))
                    city, district, road1, road2, note = split_result

                    df_valid.at[i, 'city'] = city
                    df_valid.at[i, 'district'] = district
                    df_valid.at[i, 'road_1'] = road1
                    df_valid.at[i, 'road_2'] = road2
                    df_valid.at[i, 'note'] = note
                    df_valid.at[i, 'camera_type'] = row['camera_type']
                    if pd.isna(row['speed']):
                        df_valid.at[i, 'speed'] = 0
                    else:
                        df_valid.at[i, 'speed'] = row['speed']
                    if pd.isna(row['type']):
                        df_valid.at[i, 'type'] = row['Source_File']
                    else:
                        df_valid.at[i, 'type'] = row['type']
                    # df_valid.at[i, 'Source_File'] = row['Source_File']
                    # print(df_valid.at[i, 'type'])
                else:
                    df_invalid = pd.concat([df_invalid, row.to_frame().T], ignore_index=True)

            output_file_path = os.path.join(output_path, filename)
            df_valid.to_excel(output_file_path, index=False)
            df_invalid.to_excel(os.path.join(invalid_path, filename), index=False)
            print(f"解析結果已保存到 {output_file_path}")
            print(f"不符合解析條件的數據已保存到 {os.path.join(invalid_path, filename)}")


if __name__ == "__main__":
    input_path = os.path.join(project_path, r"split_data\input")
    output_path = os.path.join(project_path, r"split_data\output")
    invalid_path = os.path.join(project_path, r"split_data\invalid")
    data_process1()
    data_process2()
    mio_data_process()
    main(input_path, output_path, invalid_path)
    data_process3()

