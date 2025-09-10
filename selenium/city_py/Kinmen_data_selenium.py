import time
from selenium import webdriver
from selenium.webdriver.common.by import By


# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\Kinmen"

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
    driver.get("https://kpb.kinmen.gov.tw/cp.aspx?n=253754F275EEBE90")
    pdf_links = driver.find_elements(By.CLASS_NAME, 'pdf')
    if len(pdf_links) >= 2:
        # 點擊第二個元素（索引從 0 開始，所以第二個元素是索引 1）
        pdf_links[1].click()
    else:
        pdf_links[0].click()
    # input("Press Enter to close the browser...")
    time.sleep(3)
finally:
    driver.quit()
