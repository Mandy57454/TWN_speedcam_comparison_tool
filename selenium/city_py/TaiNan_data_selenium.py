import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword = {
    '臺南市政府警察局固定式交通違規照相路段設置一覽表.pdf',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\TaiNan"

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
        driver.get("https://www.tnpd.gov.tw/Article/71d16651-5929-91c8-8a06-039fc3dbee6c")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()

        # input("Press Enter to close the browser...")
        time.sleep(3)
finally:
    driver.quit()
