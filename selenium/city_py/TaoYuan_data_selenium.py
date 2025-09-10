import time
from selenium import webdriver
from selenium.webdriver.common.by import By

keyword = {
    '113年桃園市政府警察局交通警察大隊科技執法設備設置及移動式地點一覽表',
}

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\TaoYuan"

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
        driver.get("https://traffic2.tyhp.gov.tw/index.php?catid=77#gsc.tab=0")
        element = driver.find_element(By.PARTIAL_LINK_TEXT, keyword)
        element.click()
        pdf_links = driver.find_elements(By.PARTIAL_LINK_TEXT, 'pdf')
        index = 0
        for index_s in pdf_links:
            pdf_links[index].click()
            index = index + 1
        time.sleep(3)
finally:
    driver.quit()
