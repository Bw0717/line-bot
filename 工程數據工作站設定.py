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
from utils.pd import read_excel_and_group_by_hierarchy


file_path = r"G:/MES自動化/utils/自動匯入表單.xlsx"
grouped_data = read_excel_and_group_by_hierarchy(file_path)
print(grouped_data)
error_messages = []

options = Options()
options.add_argument('--ignore-certificate-errors')  
options.add_argument('--disable-web-security')      


input_acc = "14574"


driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)
driver.maximize_window()
wait = WebDriverWait(driver, 10)  


try:

    driver.get(f"http://cimes.seec.com.tw/{dele_type}/Security/CimesUserLogin.aspx")

    # 
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input#UserName"))
    )
    password_input = driver.find_element(By.CSS_SELECTOR, "input#Password")
    login_button = driver.find_element(By.ID, "LoginButton")

    # 
    username_input.send_keys(input_acc)
    password_input.send_keys(input_acc)
    login_button.click()
    all_windows = driver.window_handles

    for window in driver.window_handles:
        if window != all_windows:
            driver.switch_to.window(window)
            break
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "TestMenu"))
    )

    # 
    element.click()
    driver.switch_to.frame("ifmMenu")
    element3 = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#navFRE.ctgr"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(element3).click().perform()
    element4 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPFRE"][cimesclass="EDC"]')))
    element4.click()

    element5 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#caption')))
    element5.click()

    driver.switch_to.default_content()
    iframe = driver.find_element(By.CLASS_NAME, "win1")
    driver.switch_to.frame(iframe)


    dropdown = driver.find_element(By.ID, "ddlProduct")
    select = Select(dropdown)
    select.select_by_value("ALL")
    time.sleep(1)
    for product, stations in grouped_data.items():
        input_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ttbDeviceFilter"))
        )
        input_field.clear()
        input_field.send_keys(product)


        element6 = driver.find_element(By.CSS_SELECTOR, '#btnRefresh[name="btnRefresh"]')
        actions2 = ActionChains(driver)

        actions2.move_to_element(element6).click().perform()
        time.sleep(1)
        for station, parameters in stations.items():
        #第二個For 遍歷工作站
            dropdown = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "ddlOperation"))
            )

            select = Select(dropdown)
            select.select_by_value(station)
            time.sleep(1)

            add_button = driver.find_element(By.ID, "btnAdd")
            actions2 = ActionChains(driver)
            actions2.move_to_element(add_button).click().perform()

        
            time.sleep(1)
            second_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "EDCOperSetPara.aspx")]')  # 用 xpath 定位第二層 iframe
            driver.switch_to.frame(second_iframe)  # 切換到第二層 iframe

            try:
                ddl_rule_name_element = driver.find_element(By.ID, "ddlRuleName")
                select = Select(ddl_rule_name_element)
                select.select_by_index(1)
            except:
                pass
            time.sleep(1)
            ddl_type_element = driver.find_element(By.ID, "ddlType")
            select = Select(ddl_type_element)
            select.select_by_index(1)
            time.sleep(1.5)
            dropdown_element = Select(driver.find_element(By.ID, "ddlCorelationOper"))
            dropdown_element.select_by_visible_text(station)

            #刪除動作
            if check == 1:  # 確認檢查條件是否成立
                while True:
                    elements = driver.find_elements(By.CSS_SELECTOR, ".CSGridEditButton")
                    if elements:
                        elements[0].click()
                        enable_time_input = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[@id='PanelParameterInfo']//input[@id='ttbEnableTime']"))
                        )                        
                        ActionChains(driver).move_to_element(enable_time_input).perform()
                        value = enable_time_input.get_attribute('value')
                        #if enable_time_input.is_enabled() and enable_time_input.is_displayed():
                        #    print("元素是可點擊的！")
                        #else:
                        #    print("元素無法點擊！")


                        ttbDisableTime = wait.until(EC.element_to_be_clickable((By.ID, "ttbDisableTime")))
                        ttbDisableTime.clear()
                        time.sleep(0.5)
                        ttbDisableTime.send_keys(value)
                        ttbDisableTime.send_keys(Keys.RETURN)
                        time.sleep(0.5)
                        ok_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "btnOKParameter"))
                        )                
                        ok_button.click()
                        time.sleep(1)
                        save_button5 = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "btnSave"))
                        )                
                        save_button5.click()
                        time.sleep(1)
                    else:
                        break  
            else:
                pass


        #第三個For，遍歷參數
            for parameter in parameters:
                time.sleep(1)
                driver.execute_script("document.querySelector('#btnAdd').click();")
                button11 = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnQueryParameter"))
                )
                driver.execute_script("arguments[0].click();", button11)

                input_element = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "ttbFilter"))
                )
                input_element.clear()
                time.sleep(1)
                input_element2 = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "ttbFilter"))
                )
                input_element2.click
                input_element2.send_keys(parameter["參數名稱"])
                time.sleep(0.5)
                input_element2.send_keys(Keys.RETURN)
                time.sleep(5)
                try:
                    element_by_css = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "a[href=\"javascript:__doPostBack('gvFilterItems','Select$0')\"]"))
                    )
                    driver.execute_script("arguments[0].click();", element_by_css)
                    time.sleep(0.5)
                    input_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "ttbOperSeq"))
                    )
                    input_element.clear()
                    input_element.send_keys(parameter["順序"])

                    dropdown3 = driver.find_element(By.ID, "ddlOperCritical")
                    select3 = Select(dropdown3)
                    select3.select_by_value(parameter["重要性"])

                    datetime_input2 = driver.find_element(By.ID, "ttbEnableTime")
                    datetime_input2.clear()
                    datetime_input2.send_keys("2020/12/25 17:42:12")
                    datetime_input2.send_keys(Keys.RETURN)
                    time.sleep(1)
                    btnOKParameter = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "btnOKParameter"))
                    )
                    btnOKParameter.click()
                except:
                    error_messages.append("找不到" + parameter["參數名稱"])
                    btnClose = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "btnClose"))
                    )
                    btnClose.click()
                    time.sleep(1)
                    continue

                    

                try:
                    time.sleep(4)
                    error_message = WebDriverWait(driver, 2).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "span#lblParaMsg.CSMust"))
                    )
                    error_text = error_message.text
                    error_messages.append(error_message.text)

                    time.sleep(1)
                    btnCCParameter = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.ID, "btnCloseParameter"))
                    )
                    btnCCParameter.click()
                except:
                    pass

                time.sleep(3)

            #下一個參數，直到結束
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btnSave"))
            )
            save_button.click()
            time.sleep(2)

            exit_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btnExit"))
            )
            exit_button.click()

            alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert.accept()

            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            time.sleep(1)
    #下一個工作站，直到結束
    #下一個件號，直到結束


    
    
   #add_button2 = WebDriverWait(driver, 10).until(
   #     EC.element_to_be_clickable((By.ID, "btnAdd"))
   # )
   # add_button2.click()
   # time.sleep(1)
   # second_iframe = driver.find_element(By.XPATH, '//iframe[contains(@src, "EDCOperSetPara.aspx")]')  # 用 xpath 定位第二層 iframe
   # driver.switch_to.frame(second_iframe)  # 切換到第二層 iframe
   # ddl_rule_name_element = driver.find_element(By.ID, "ddlRuleName")
   # select = Select(ddl_rule_name_element)
   # select.select_by_index(1)

    with open("error_messages.txt", "w", encoding="utf-8") as file:
        for message in error_messages:
            file.write(message + "\n")
    time.sleep(5)
finally:

    driver.quit()