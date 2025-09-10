import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from city_processor import process_taipei, process_TaoYuan, process_KeeLung, process_ChiaYi_sh, process_HsinChu_web
from city_processor import process_MiaoLi_YunLin_NanTou, process_TaiChung, process_ChiaYi_YiLan, process_TaiTung
from city_processor import process_Kaohsiung, process_PengHu, process_TaiNan, process_HsinChu_pdf
from city_processor import process_ChangHua, process_HuaLian_KinMen
from pdf_to_excel import pdf_2_excel

city_url_keywords = {
    'taipei': {
        'url': 'https://td.police.gov.taipei/Content_List.aspx?n=6495BB8B3BA7248D',
        'keywords_index_s': {
            '固定測速桿地點': '1',
            '移動式測速照相地點': '0',
            '固定式科技執法地點': '1',
        }
    },
    'TaoYuan': {
        'url': 'https://traffic2.tyhp.gov.tw/index.php?catid=77#gsc.tab=0',
        'keywords_index_s': {
            '113年桃園市政府警察局交通警察大隊科技執法設備設置及移動式地點一覽表': {
                '1', '2',
            }
        }
    },
    'KeeLung': {
        'url': 'https://www.klg.gov.tw/cht/index.php?code=list&ids=937',
        'keywords_index_s': {
            '闖紅燈及測速自動照相設備設置地點': '0',
            '本局現有「固定式科技執法設備設置地點一覽表」': '1',
            '本市租賃固定式科技執法設備設置地點一覽表': '0',
            '固定式科技執法區間平均速率自動偵測照相系統設置地點': '0',
            '移動式測速照相地點': '0',
        }
    },
    'HsinChu_pdf': {
        'url': 'https://www.2spc.npa.gov.tw/ch/app/data/view?module=wg020&id=2047&serno=bec74d67-823d-4e5c-bd75-c0cd9dd9f1b1',
    },
    'HsinChu_web': {
        'url': 'https://tra.hccp.gov.tw/pages/camera',
        'city': 'HsinChu_web',
    },
    'MiaoLi': {
        'url': 'https://www.mpb.gov.tw/sub/latestevent/index?Parser=9,27,521,504',
        'keywords_index_s': {
            '公告本局「固定式闖紅燈、測速、跨越雙黃線、區間平均速率自動偵測照相設備、路口多功能科技執法系統及違規停車自動偵測系統」設置地點一覽表': '1',
        }
    },
    'TaiChung': {
        'url': 'https://www.police.taichung.gov.tw/traffic/home.jsp?id=55&parentpath=0,5,53',
        'keywords_index_s': {
            '臺中市政府警察局「科學儀器執法設備」固定式及移動式取締地點一覽表': '1',
        }
    },
    'YunLin': {
        'url': 'https://www.ylhpb.gov.tw/latestevent/index.aspx?Parser=22,3,31,,,,,,,,0',
        'keywords_index_s': {
            '雲林縣警察局固定式違規照相設備及區間平均速率自動偵測照相系統設置地點一覽表': '0',
        }
    },
    'ChiaYi': {
        'url': 'https://www.cypd.gov.tw/Tpb/Directory/a18d2191-3ec0-7fbc-5c0b-b1de3ab62f23',
        'keywords_keyWords': {
            '固定桿測速照相地點': '嘉義縣警察局固定照相桿地點一覽表',
            '非固定式測速照相設置地點': '分局非固定式測速照相設置地點',
        }
    },
    'NanTou': {
        'url': 'https://www.ncpd.gov.tw/sub/latestevent/index.aspx?Parser=9,30,755,736',
        'keywords_index_s': {
            '公告本局113年轄區現有及新增「科學儀器執法設備」設置地點一覽表': '0',
        }
    },
    'TaiNan': {
        'url': 'https://www.tnpd.gov.tw/Article/71d16651-5929-91c8-8a06-039fc3dbee6c',
        'keywords': {
            '臺南市政府警察局固定式交通違規照相路段設置一覽表.pdf',
        }
    },
    'TaiTung': {
        'url': 'https://www.ttcpb.gov.tw/chinese/home.jsp?serno=201105170134&contlink=ap/traffic_view.jsp&dataserno=202109060001',
    },
    'HuaLian': {
        'url': 'https://www.hlpb.gov.tw/contents/80?node=49&trans=1',
    },
    'YiLan': {
        'url': 'https://www.ilcpb.gov.tw/Message?itemid=653&mid=5651',
        'keywords_keyWords': {
            '本局固定式科學儀器執法設備設置地點一覽表': '設置地點一覽表.pdf',
        }
    },
    'KinMen': {
        'url': 'https://kpb.kinmen.gov.tw/cp.aspx?n=253754F275EEBE90',
    },
    'Kaohsiung': {
        'url': 'https://kcpd.kcg.gov.tw/cp.aspx?n=693052840FE00C08&s=D36FF802DE9B88C4',
        'keywords_keyWords': {
            '固定式違規照相科技執法設備設置地點': '固定式違規照相科技執法設備設置地點',
            '區間平均速率執法設備設置地點': '區間平均速率執法設備設置地點',
            '路口科技執法監測系統設置地點': '路口科技執法監測系統設置地點',
            '不停讓行人科技執法監測系統設置地點': '不停讓行人科技執法監測系統設置地點',
            '違規停車科技執法設置地點': '「違規停車」監測系統設置地點',
            '租賃式闖紅燈科技執法照相設備設置地點': '租賃式闖紅燈科技執法照相設備設置地點',
            '交通局建置科技執法設備設置地點': '交通局建置科技執法設備設置地點',
            '移動式測速照相重點路段': '高雄市政府警察局移動式測速照相重點路段',
            '捷運局輕軌沿線建置科技執法設備設置地點': '捷運局輕軌沿線建置科技執法設備設置地點',
        },
    },
    'ChiaYi_sh': {
        'url': 'https://www.ccpb.gov.tw/content/?parent_id=10081###',
        'city': 'ChiaYi_sh',
    },
    'PengHu': {
        'url': 'https://www.phpb.gov.tw/home.jsp?intpage=1&id=131&qptdate=&qdldate=&keyword=%E8%AB%8B%E8%BC%B8%E5%85%A5%E9%97%9C%E9%8D%B5%E5%AD%97&pagenum=1&pagesize=15',
    },
    'ChangHua': {
        'url': 'https://www.chpb.gov.tw/FileList/C004200?SubCategoryID=b41fcf60-8d2b-4a7d-8a58-3e5e7e691e2b',
    },
}
city_function_map = {
        'taipei': process_taipei,
        'TaoYuan': process_TaoYuan,
        'KeeLung': process_KeeLung,
        'HsinChu_pdf': process_HsinChu_pdf,
        'HsinChu_web': process_HsinChu_web,
        'MiaoLi': process_MiaoLi_YunLin_NanTou,
        'TaiChung': process_TaiChung,
        'YunLin': process_MiaoLi_YunLin_NanTou,
        'ChiaYi': process_ChiaYi_YiLan,
        'NanTou': process_MiaoLi_YunLin_NanTou,
        'TaiNan': process_TaiNan,
        'TaiTung': process_TaiTung,
        'HuaLian': process_HuaLian_KinMen,
        'YiLan': process_ChiaYi_YiLan,
        'KinMen': process_HuaLian_KinMen,
        'Kaohsiung': process_Kaohsiung,
        'ChiaYi_sh': process_ChiaYi_sh,
        'PengHu': process_PengHu,
        'ChangHua': process_ChangHua,
    }


def clear_folder(folder_path):
    # 清空文件夹中的所有内容
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # 删除文件或符号链接
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # 删除文件夹
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')


def main():
    for city, info in city_url_keywords.items():
        if city == 'taipei':
            # 存檔路徑
            download_path = os.path.join(r"C:\Users\mandy.chang\PycharmProjects\selenium\city_data", city)
            if os.path.exists(download_path):
                clear_folder(download_path)
            else:
                os.makedirs(download_path)

            service = Service(r'D:\Taiwan speedcam\selenium\chromedriver-win64\chromedriver.exe')
            # 配置 Chrome 下載選項
            chrome_options = webdriver.ChromeOptions()
            prefs = {
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": True  # 這樣就不會在 Chrome 內嵌閱讀 PDF，而是直接下載
            }
            chrome_options.add_experimental_option("prefs", prefs)
            # 設置driver
            driver = webdriver.Chrome(options=chrome_options)
            city_function_map[city](driver, info)
            print(city, "data download complete")

            # pdf_2_excel(chrome_options, service, city)
    print("all data download complete")


if __name__ == "__main__":
    main()
