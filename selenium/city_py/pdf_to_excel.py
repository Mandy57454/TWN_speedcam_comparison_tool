from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pyautogui
import os


def open_f12():
    # 打開開發者工具
    pyautogui.press('f12')
    time.sleep(10)
    more_tab_position = (685, 252)
    application_tab_position = (721, 369)
    storage_tab_position = (455, 350)
    pyautogui.click(more_tab_position)
    time.sleep(1)
    pyautogui.click(application_tab_position)
    time.sleep(1)
    pyautogui.click(storage_tab_position)
    time.sleep(1)


def cookie_reset():
    clear_tab_position = (672, 626)
    pyautogui.click(clear_tab_position)
    time.sleep(2)
    pyautogui.press('f5')
    time.sleep(2)
    pyautogui.press('enter')


def wait_for_page_load(driver, timeout=10):
    start_time = time.time()
    while time.time() - start_time < timeout:
        # 检查页面的 readyState
        ready_state = driver.execute_script("return document.readyState")
        if ready_state == "complete":
            return True
        else:
            time.sleep(0.5)
    return False


def pdf_2_excel(chrome_options, service, city):
    folder_path = os.path.join(r"C:\Users\mandy.chang\PycharmProjects\selenium\city_data", city)
    # window_size = driver.get_window_size()
    # print(f"Current window size: width = {window_size['width']}, height = {window_size['height']}")

    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            driver = webdriver.Chrome(options=chrome_options, service=service)
            try:
                # 打开网页
                driver.get('https://www.adobe.com/tw/acrobat/online/pdf-to-excel.html')
                # driver.set_window_size(945, 1012)
                # 等待页面加载
                wait_for_page_load(driver)
                open_f12()
                file_path_ = os.path.join(folder_path, filename)
                cookie_reset()
                wait_for_page_load(driver)
                file_upload_element = driver.find_element(By.ID, 'fileInput')
                # 使用 send_keys 将文件路径发送到文件上传控件
                file_upload_element.send_keys(file_path_)
                time.sleep(35)
                download_tab_position = (179, 521)
                pyautogui.click(download_tab_position)
                time.sleep(5)
            except Exception as e:
                print(f"Error processing {filename}: {e}")
            finally:
                driver.quit()

# 定位鼠標位置
# current_position = pyautogui.position()
# print("Current mouse position:", current_position)

