import os
from selenium import webdriver
import time
import openpyxl
from selenium.webdriver.common.by import By

city_url_keywords = {
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
}

for city, info in city_url_keywords.items():
    # 存檔路徑
    download_path = os.path.join(r"C:\Users\mandy.chang\PycharmProjects\selenium\city_data", city)
    if os.path.exists(download_path):
        print("folder exists")
    else:
        os.makedirs(download_path)

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

    url = info['url']
    keywords_keyWords = info['keywords_keyWords']
    try:
        for keyword, keyWord in keywords_keyWords.items():
            driver.get(url)
            element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
            element.click()
            table = driver.find_element(By.XPATH, "//table[caption[contains(text(), keyWord)]]")
            rows = table.find_elements(By.TAG_NAME, "tr")
            # 創建一個新的 Excel 工作簿
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = keyword
            for row_index, row in enumerate(rows, start=1):
                # 获取行中的所有单元格
                cells = row.find_elements(By.TAG_NAME, "td")
                for cell_index, cell in enumerate(cells, start=1):
                    # 將單元格的文本內容寫入對應的 Excel 單元格
                    sheet.cell(row=row_index, column=cell_index, value=cell.text)
            # 保存 Excel 文件
            filename = f"{keyword}.xlsx"
            workbook.save(filename)

    finally:
        driver.quit()
