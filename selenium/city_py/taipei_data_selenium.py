import time
from selenium.webdriver.common.by import By


def taipei(driver, info):
    url = info['url']
    keywords_index_s = info['keywords_index_s']
    try:
        for word, index_s in keywords_index_s.items():
            driver.get(url)
            element = driver.find_element(By.PARTIAL_LINK_TEXT, word)
            element.click()
            pdf_links = driver.find_elements(By.CLASS_NAME, 'pdf')

            for index in index_s:
                pdf_links[int(index)].click()
            time.sleep(3)
    finally:
        driver.quit()
