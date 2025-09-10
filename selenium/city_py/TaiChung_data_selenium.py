import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword_index = {
    '臺中市政府警察局「科學儀器執法設備」固定式及移動式取締地點一覽表': '1',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\TaiChung"

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
        driver.get("https://www.police.taichung.gov.tw/traffic/home.jsp?id=55&parentpath=0,5,53")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_links = driver.find_elements(By.XPATH, "//img[@alt='PDF']")
        pdf_links[int(index)].click()
        time.sleep(3)
finally:
    driver.quit()
