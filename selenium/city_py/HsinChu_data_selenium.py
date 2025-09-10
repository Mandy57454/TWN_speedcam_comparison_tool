import time
from selenium import webdriver
from selenium.webdriver.common.by import By


# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\HsinChu"

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
    driver.get("https://www.2spc.npa.gov.tw/ch/app/data/view?module=wg020&id=2047&serno=bec74d67-823d-4e5c-bd75-c0cd9dd9f1b1")
    element = driver.find_element(By.XPATH, "//img[@src='/static/images/file/pdf.png']")
    element.click()
    # input("Press Enter to close the browser...")
    time.sleep(3)
finally:
    driver.quit()

# <img src="/static/images/file/pdf.png" alt="保二總隊科技執法設置情形一覽表-新1120113">
