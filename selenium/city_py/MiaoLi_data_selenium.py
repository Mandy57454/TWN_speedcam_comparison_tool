import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword_index = {
    '公告本局「固定式闖紅燈、測速、跨越雙黃線、區間平均速率自動偵測照相設備、路口多功能科技執法系統及違規停車自動偵測系統」設置地點一覽表': '1',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\MiaoLi"

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
        driver.get("https://www.mpb.gov.tw/sub/latestevent/index?Parser=9,27,521,504")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_links = driver.find_elements(By.XPATH, "//img[@src='/Content/img/d06.gif']")
        pdf_links[int(index)].click()
        time.sleep(3)
finally:
    driver.quit()

# <img src="/Content/img/d06.gif" alt="下載PDF檔案(1130701公告地點一覽表.pdf)_另開視窗" width="16" height="16" border="0" align="absmiddle">
