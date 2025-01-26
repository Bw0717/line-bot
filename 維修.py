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
input_QR = "31110AKC8E000--KYD4020001540020250116---"
try:
    driver.get("http://cimes.seec.com.tw/ARG/CimesDesktop.aspx")
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
    element7 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#divOuter a.CSLot")))
    driver.execute_script("arguments[0].click();", element7)
    actions.move_to_element(element7).click().perform()
    Niframes = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(Niframes[0])
    a_element2 = driver.find_element(By.XPATH, "//a[span[text()='歷史記錄']]")
    driver.execute_script("arguments[0].click();", a_element2)
    btnHisQuery = wait.until(EC.presence_of_element_located((By.ID, "btnHisQuery")))
    driver.execute_script("arguments[0].click();", btnHisQuery)
    buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='button' and starts-with(@href, 'javascript:__doPostBack')]")))
    buttons[-1].click()
    time.sleep(5)
    try:
        tbody = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'table#gvHist > tbody'))
        )
        rows = tbody.find_elements(By.CSS_SELECTOR, 'tr')
        print(len(rows))
        last_row = rows[-1]
        last_column = last_row.find_elements(By.CSS_SELECTOR, 'td')[4].text  
        print("最後一筆工作站:", last_column)

        """target_workstation = "ACGF飛輪線邊站"


        for row in rows:

            workstation = row.find_elements(By.CSS_SELECTOR, 'td')[4].text
            print(workstation)
            if target_workstation in workstation:
                print(f"找到匹配的工作站: {workstation}")

                
        else:
            print("找不到匹配的工作站")"""
    except:
        pass
    btnExit = wait.until(EC.presence_of_element_located((By.ID, "btnExit")))
    driver.execute_script("arguments[0].click();", btnExit)
    time.sleep(1)
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
    element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="WipRule"]')))
    element4.click()
    button2 = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@title, '維修(客)')]")))
    actions.move_to_element(button2).click().perform()
    driver.switch_to.default_content()
    Giframes = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(Giframes[1])
    input1 = wait.until(EC.presence_of_element_located((By.ID, "ttbLot")))
    input1.send_keys(input_QR)
    input1.send_keys(Keys.RETURN)
    time.sleep(1.5)
    select_list = wait.until(EC.presence_of_element_located((By.ID, "ddlDutyUnit")))
    select = Select(select_list)
    select.select_by_visible_text("AR1")
    select_list2 = wait.until(EC.presence_of_element_located((By.ID, "ddlReturnOperation")))
    select = Select(select_list2)
    select.select_by_visible_text(last_column)
    input2 = wait.until(EC.presence_of_element_located((By.ID, "ttbRepairTime")))
    input2.send_keys("0")
    finish_btnOK = wait.until(EC.presence_of_element_located((By.ID, "btnOK")))
    finish_btnOK.click()
    time.sleep(2)    
finally:
    driver.quit()
