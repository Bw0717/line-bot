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
options = Options()
options.add_argument('--ignore-certificate-errors')  
options.add_argument('--disable-web-security')      
input_acc = "14574"
driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)
driver.maximize_window()
actions2 = ActionChains(driver)
wait = WebDriverWait(driver, 10)  
wait2 = WebDriverWait(driver, 20)  
data_input = "277959"
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

    # 
    element_menu.click()
    driver.switch_to.frame("ifmMenu")
    op_element = wait.until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#navRULE.ctgr_hov"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(op_element).click().perform()

    element4 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPRULE"][cimesclass="CustRule"]')))
    element4.click()

    button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="日工單開立"]')))
    actions.move_to_element(button).click().perform()
    
    time.sleep(3)
    driver.switch_to.default_content()
    iframe = driver.find_element(By.CLASS_NAME, "win1")
    driver.switch_to.frame(iframe)

    input_element3 = wait.until(EC.presence_of_element_located((By.ID, "CimesInputBox")))
    keyboard.press_and_release('ctrl+-')

    input_element3.send_keys(data_input)
    input_element3.send_keys(Keys.RETURN)
    time.sleep(2)
    enable_input = wait.until(EC.presence_of_element_located((By.ID, "ttbUnCreateQty")))
    value = enable_input.get_attribute('value')
    input_element4 = wait.until(EC.presence_of_element_located((By.ID, "ttbDayWOQty")))
    input_element4.send_keys(value)
    add_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "btnAdd"))
    )                
    add_button.click()

    ok_button = wait.until(
    EC.element_to_be_clickable((By.ID, "btnOK"))
    )             
    actions.move_to_element(ok_button).click().perform()
    time.sleep(3)    


finally:
    driver.quit()

