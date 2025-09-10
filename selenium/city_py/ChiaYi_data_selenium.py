import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword_keyWords = {
    '固定桿測速照相地點': '嘉義縣警察局固定照相桿地點一覽表',
    '非固定式測速照相設置地點': '分局非固定式測速照相設置地點',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\ChiaYi"

# 配置 Chrome 下載選項
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True  # 這樣就不會在 Chrome 內嵌閱讀 PDF，而是直接下載
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

try:
    for keyword, keyWords in keyword_keyWords.items():
        driver.get("https://www.cypd.gov.tw/Tpb/Directory/a18d2191-3ec0-7fbc-5c0b-b1de3ab62f23")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_links = driver.find_elements(By.PARTIAL_LINK_TEXT, keyWords)

        index = 0
        for index_s in pdf_links:
            pdf_links[index].click()
            index = index+1

        # input("Press Enter to close the browser...")
        time.sleep(3)
finally:
    driver.quit()
