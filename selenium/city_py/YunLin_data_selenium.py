import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword = {
    '雲林縣警察局固定式違規照相設備及區間平均速率自動偵測照相系統設置地點一覽表',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\YunLin"

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
    for keyword in keyword:
        driver.get("https://www.ylhpb.gov.tw/latestevent/index.aspx?Parser=22,3,31,,,,,,,,0")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_link = driver.find_elements(By.XPATH, "//img[@src='/Content/img/d06.gif']")
        pdf_link[0].click()
        time.sleep(3)
finally:
    driver.quit()
