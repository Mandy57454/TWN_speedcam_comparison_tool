import os
import shutil
import traceback
from selenium import webdriver
from city_processor import process_taipei, process_TaoYuan, process_KeeLung, process_ChiaYi_sh, process_HsinChu_web
from city_processor import process_MiaoLi_YunLin, process_TaiChung, process_ChiaYi_YiLan, process_TaiTung, process_NanTou
from city_processor import process_Kaohsiung, process_PengHu, process_TaiNan, process_Science_Park, process_HsinChu
from city_processor import process_ChangHua, process_HuaLian_KinMen, process_newtaipei
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException
# from pdf_to_excel import pdf_2_excel

# 動態取得專案路徑
project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# cities url and keywords
city_url_keywords = {
    'taipei': {
        'url': 'https://td.police.gov.taipei/Content_List.aspx?n=6495BB8B3BA7248D',
        'keywords_index_s': {
            '固定測速桿地點': [],
            '移動式測速照相地點': [],
            '固定式科技執法地點': [],
        }
    },
    'TaoYuan': {
        'url': 'https://traffic2.tyhp.gov.tw/index.php?catid=77#gsc.tab=0',
        'keywords_index_s': {
            '115年桃園市政府警察局交通警察大隊科技執法設備設置及移動式地點一覽表': [],  # 空陣列，下載所有 PDF
        }
    },
    'KeeLung': {
        'url': 'https://www.klg.gov.tw/cht/index.php?code=list&ids=937',
        'keywords': {
            '闖紅燈及測速自動照相設備設置地點',
            '本局現有「固定式科技執法設備設置地點一覽表」',
            '本市租賃固定式科技執法設備設置地點一覽表',
            '固定式科技執法區間平均速率自動偵測照相系統設置地點',
            '移動式測速照相地點',
        }
    },
    'Science_Park': {
        'url': 'https://www.2spc.npa.gov.tw/ch/app/data/view?module=wg020&id=2047&serno=bec74d67-823d-4e5c-bd75-c0cd9dd9f1b1',
    },
    'HsinChu_web': {
        'url': 'https://tra.hccp.gov.tw/pages/camera',
        'city': 'HsinChu_web',
    },
    'MiaoLi': {
        'url': 'https://www.mpb.gov.tw/sub/latestevent/index?Parser=9,27,521,504',
        'keywords_index_s': {
            '公告本局「固定式闖紅燈、測速、區間平均速率自動偵測照相設備、路口多功能科技執法系統及違規停車自動偵測系統」設置地點一覽表': {'1', '2'},
        }
    },
    'TaiChung': {
        'url': 'https://www.police.taichung.gov.tw/traffic/home.jsp?id=55&parentpath=0,5,53',
        'keywords_index_s': {
            '臺中市政府警察局「科學儀器執法設備」固定式及移動式取締地點一覽表': {'0', '1', '3'},
        }
    },
    'YunLin': {
        'url': 'https://www.ylhpb.gov.tw/latestevent/index.aspx?Parser=22,3,31,,,,,,,,0',
        'keywords_index_s': {
            '雲林縣警察局固定式違規照相設備及區間平均速率自動偵測照相系統設置地點一覽表': '0',
        }
    },
    'newtaipei': {
        'url': 'https://www.traffic.police.ntpc.gov.tw/lp-3313-27.html',
        'fixed': {'新北市「固定式科學儀器執法設備設置地點」一覽表及公告': []},
        'mobile': '移動式科學儀器執法設備設置地點'
    },
    'HsinChu': {
        'url': 'https://www.hchpb.gov.tw/Tw/Home/Index?SiteID=416763f8-a1f7-48fd-bfbf-9327913efad7',
        'keywords': {
            '固定式科學儀器執法設備設置地點',
            '移動式測速照相執行地點一覽表'
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
        'url': 'https://www.ncpd.gov.tw/sub/table/index.aspx?Parser=5,30,1147,737,,,,,,,0',
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
            '本局固定式 ': '設置地點一覽表',
        }
    },
    'KinMen': {
        'url': 'https://kpb.kinmen.gov.tw/cp.aspx?n=253754F275EEBE90',
    },
    'Kaohsiung': {
        'url': 'https://kcpd.kcg.gov.tw/cp.aspx?n=693052840FE00C08',
        'keywords_keyWords': {
            '固定式違規照相科技執法設備設置地點': '固定式違規照相科技執法設備設置地點',
            '區間平均速率執法設備設置地點': '區間平均速率執法設備設置地點',
            '路口科技執法監測系統設置地點': '路口科技執法監測系統設置地點',
            '不停讓行人科技執法監測系統設置地點': '不停讓行人科技執法監測系統設置地點',
            '限制車種及違規停車科技執法設置地點': '違規停車',
            '交通局建置科技執法設備設置地點': '交通局建置科技執法設備設置地點',
            '移動式測速照相': '移動式測速照相',
            '捷運局輕軌沿線建置科技執法設備設置地點': '捷運局輕軌沿線建置科技執法設備設置地點',
            '租賃式車不停讓行人科技執法監測系統設置地點': '租賃式車不停讓行人科技執法監測系統設置地點',
        },
    },
    'ChiaYi_sh': {
        'url': 'https://www.ccpb.gov.tw/news/?division=17&keywords=&mode=search&type_id=10326&parent_id=10321',
        'keywords_keyWords': {
            '公告本局轄區「固定式科學儀器執法設備」設置地點一覽表': '固定式科學儀器執法設備設置地點一覽表.pdf',
        }
    },
    'PengHu': {
        'url': 'https://www.phpb.gov.tw/home.jsp?intpage=1&id=131&qptdate=&qdldate=&keyword=%E8%AB%8B%E8%BC%B8%E5%85%A5%E9%97%9C%E9%8D%B5%E5%AD%97&pagenum=1&pagesize=15',
    },
    'ChangHua': {
        'url': 'https://www.chpb.gov.tw/FileList/C004200?SubCategoryID=b41fcf60-8d2b-4a7d-8a58-3e5e7e691e2b',
    },
}
# cities function
city_function_map = {
        'taipei': process_taipei,
        'TaoYuan': process_TaoYuan,
        'KeeLung': process_KeeLung,
        'Science_Park': process_Science_Park,
        'HsinChu_web': process_HsinChu_web,
        'HsinChu': process_HsinChu,
        'MiaoLi': process_MiaoLi_YunLin,
        'TaiChung': process_TaiChung,
        'YunLin': process_MiaoLi_YunLin,
        'ChiaYi': process_ChiaYi_YiLan,
        'NanTou': process_NanTou,
        'TaiNan': process_TaiNan,
        'TaiTung': process_TaiTung,
        'HuaLian': process_HuaLian_KinMen,
        'YiLan': process_ChiaYi_YiLan,
        'KinMen': process_HuaLian_KinMen,
        'Kaohsiung': process_Kaohsiung,
        'ChiaYi_sh': process_ChiaYi_sh,
        'PengHu': process_PengHu,
        'ChangHua': process_ChangHua,
        'newtaipei': process_newtaipei
    }


# check folder exists and clear folder all content
def clear_folder(folder_path):
    # loop through all files in the folder
    for filename in os.listdir(folder_path):
        # construct the full file path
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # 删除文件或符號連結
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # 删除文件夾
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')


def main(debug_city=None):
    """
    主函數
    
    Args:
        debug_city (str, optional): 如果指定，只處理該縣市。例如: 'TaoYuan', 'taipei' 等
    """
    # 記錄成功和失敗的縣市
    successful_cities = []
    failed_cities = []
    
    # 如果指定了 debug 縣市，只處理該縣市
    if debug_city:
        if debug_city not in city_url_keywords:
            print(f"❌ 錯誤：找不到縣市 '{debug_city}'")
            print(f"可用的縣市: {', '.join(city_url_keywords.keys())}")
            return
        print(f"🔧 DEBUG 模式：只處理 {debug_city} 縣市")
        cities_to_process = {debug_city: city_url_keywords[debug_city]}
    else:
        cities_to_process = city_url_keywords
    
    for city, info in cities_to_process.items():
        driver = None
        try:
            print(f"\n開始處理 {city} 縣市...")
            
            # 存檔路徑
            download_path = os.path.join(project_path, r"selenium\city_data", city)
            # check download path exists
            if os.path.exists(download_path):
                clear_folder(download_path)
            else:
                os.makedirs(download_path)

            # setting Chrome
            chrome_options = webdriver.ChromeOptions()
            prefs = {
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": True  # 這樣就不會在 Chrome 內嵌閱讀 PDF，而是直接下載
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # 設定超時時間
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--remote-debugging-port=9222")
            
            # set driver
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(60)  # 設定頁面載入超時時間
            
            # process city data
            city_function_map[city](driver, info)
            print(f"✅ {city} 資料下載完成")
            successful_cities.append(city)
            
        except WebDriverException as e:
            print(f"❌ {city} WebDriver 錯誤: {str(e)}")
            failed_cities.append((city, f"WebDriver 錯誤: {str(e)}"))
            
        except TimeoutException as e:
            print(f"❌ {city} 頁面載入超時: {str(e)}")
            failed_cities.append((city, f"頁面載入超時: {str(e)}"))
            
        except NoSuchElementException as e:
            print(f"❌ {city} 找不到網頁元素: {str(e)}")
            failed_cities.append((city, f"找不到網頁元素: {str(e)}"))
            
        except Exception as e:
            print(f"❌ {city} 發生未預期錯誤: {str(e)}")
            print(f"錯誤詳情: {traceback.format_exc()}")
            failed_cities.append((city, f"未預期錯誤: {str(e)}"))
            
        finally:
            # 確保 driver 被正確關閉
            if driver:
                try:
                    driver.quit()
                except Exception as e:
                    print(f"⚠️ 關閉 {city} 的 WebDriver 時發生錯誤: {str(e)}")
            
            print(f"完成 {city} 的處理，繼續下一個縣市...")
    
    # 輸出處理結果摘要
    print("\n" + "="*50)
    print("資料下載處理完成摘要")
    print("="*50)
    print(f"✅ 成功處理的縣市 ({len(successful_cities)} 個):")
    for city in successful_cities:
        print(f"   - {city}")
    
    if failed_cities:
        print(f"\n❌ 處理失敗的縣市 ({len(failed_cities)} 個):")
        for city, error in failed_cities:
            print(f"   - {city}: {error}")
    else:
        print("\n🎉 所有縣市都處理成功！")
    
    print(f"\n總計: {len(successful_cities)} 成功, {len(failed_cities)} 失敗")
    print("="*50)


if __name__ == "__main__":
    import sys
    
    # 檢查是否有命令列參數指定要 debug 的縣市
    if len(sys.argv) > 1:
        debug_city = sys.argv[1]
        print(f"🚀 啟動程式，DEBUG 模式：只處理 {debug_city}")
        main(debug_city=debug_city)
    else:
        print("🚀 啟動程式，處理所有縣市")
        main()