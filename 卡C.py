from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import pandas as pd
import keyboard
from selenium.common.exceptions import StaleElementReferenceException


options = Options()
options.add_argument('--ignore-certificate-errors')  
options.add_argument('--disable-web-security')      
input_acc = "14574"
driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)
driver.maximize_window()
actions2 = ActionChains(driver)
wait = WebDriverWait(driver, 10)  
wait2 = WebDriverWait(driver, 300)  
input_QR = "8414443021786791020031600014"
try:
    driver.get("http://172.16.16.156/ARG/CimesDesktop.aspx")
    username_input = wait.until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
    )
    password_input = driver.find_element(By.CSS_SELECTOR, "input#Password")
    login_button = driver.find_element(By.ID, "LoginButton")
    username_input.send_keys(input_acc)
    password_input.send_keys(input_acc)
    login_button.click()
    all_windows = driver.window_handles
    for window in driver.window_handles:
        if window != all_windows:
            driver.switch_to.window(window)
            break
    element_menu = wait.until(
        EC.element_to_be_clickable((By.ID, "TestMenu"))
    )
    element_menu.click()
    driver.switch_to.frame("ifmMenu")
    op_element = wait.until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(op_element).click().perform()
    element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="QryProg"]')))
    element4.click()
    button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="在製品查詢"]')))
    actions.move_to_element(button).click().perform()
    driver.switch_to.default_content()
    iframe = driver.find_element(By.CLASS_NAME, "win1")
    driver.switch_to.frame(iframe)
    keyboard.press_and_release('ctrl+-')
    keyboard.press_and_release('ctrl+-')

    element5 = wait2.until(EC.element_to_be_clickable((By.XPATH, '//a[text()="批號"]')))
    element6 = driver.find_element(By.ID, "btnLotQuery")
    element6.click()
    iframe2 = wait.until(EC.presence_of_element_located((By.XPATH, '//iframe[contains(@src, "LotListFieldSelect.aspx")]')))
    driver.switch_to.frame(iframe2)
    checkbox = wait.until(EC.presence_of_element_located((By.ID, "tvn0CheckBox")))
    if checkbox.is_selected():
        checkbox.click()
    else:
        pass
    input_element3 = wait.until(EC.presence_of_element_located((By.ID, "gvFilter_ctl03_ttbValue")))
    input_element3.clear()
    input_element3.send_keys(input_QR)
    Query = wait.until(EC.presence_of_element_located((By.ID, "btnQuery")))
    Query.click()
    time.sleep(1)
    driver.switch_to.default_content()
    iframe = driver.find_element(By.CLASS_NAME, "win1")
    driver.switch_to.frame(iframe)
    elementQR = driver.find_element(By.XPATH, f'//td[text()={input_QR}]')
    elementQR.click()
    driver.switch_to.default_content()
    element_menu = wait.until(
        EC.element_to_be_clickable((By.ID, "TestMenu"))
    )
    element_menu.click()

    driver.switch_to.frame("ifmMenu")
    op_element = wait.until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(op_element).click().perform()
    element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="WipRule_Modify"]')))
    element4.click()    
    button2 = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@title, '變更工作站')]")))
    actions.move_to_element(button2).click().perform()
    time.sleep(1)
    driver.switch_to.default_content()
    Giframes = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(Giframes[1])
    select_list = wait.until(EC.presence_of_element_located((By.ID, "_ddlcsOperation")))
    select = Select(select_list)
    select.select_by_index(2)
    time.sleep(0.5)
    select_list2 = wait.until(EC.presence_of_element_located((By.ID, "_ddlcsReason")))
    select2 = Select(select_list2)
    select2.select_by_index(1)
    time.sleep(0.5)
    finish_btnOK = wait.until(EC.presence_of_element_located((By.ID, "btnOK")))
    finish_btnOK.click()
    time.sleep(5)    
finally:
    driver.quit()