import time
from selenium import webdriver
from selenium.webdriver.common.by import By

# 設定下載路徑
download_path = r"C:\Users\mandy.chang\PycharmProjects\selenium\TaiTung"

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
    driver.get("https://www.ttcpb.gov.tw/chinese/home.jsp?serno=201105170134&contlink=ap/traffic_view.jsp&dataserno=202109060001")
    element = driver.find_element(By.XPATH, "//img[@alt='pdf']")
    element.click()
    time.sleep(3)
finally:
    driver.quit()
