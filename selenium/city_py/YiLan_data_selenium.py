import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword_keyWords = {
    '本局固定式科學儀器執法設備設置地點一覽表': '設置地點一覽表.pdf',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\YiLan"

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
        driver.get("https://www.ilcpb.gov.tw/Message?itemid=653&mid=5651")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_links = driver.find_elements(By.PARTIAL_LINK_TEXT, keyWords)
        pdf_links[0].click()

        # input("Press Enter to close the browser...")
        time.sleep(3)
finally:
    driver.quit()
