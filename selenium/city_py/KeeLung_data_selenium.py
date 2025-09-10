import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword_index = {
    '闖紅燈及測速自動照相設備設置地點': '0',
    '本局現有「固定式科技執法設備設置地點一覽表」': '1',
    '本市租賃固定式科技執法設備設置地點一覽表': '0',
    '固定式科技執法區間平均速率自動偵測照相系統設置地點': '0',
    '移動式測速照相地點': '0',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\KeeLung"

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
    for keyword, index in keyword_index.items():
        driver.get("https://www.klg.gov.tw/cht/index.php?code=list&ids=937")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_links = driver.find_elements(By.CLASS_NAME, 'fileType')
        pdf_links[int(index)].click()
        # input("Press Enter to close the browser...")
        time.sleep(3)
finally:
    driver.quit()
