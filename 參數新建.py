from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from utils.pd import read_excel_and_process

# 设置文件路径
file_path = r"G:/MES自動化/utils/自動匯入表單.xlsx"

# 调用函数并获取返回的数据
data = read_excel_and_process(file_path)
edit_data = []
# 打印返回的数据
print(data)
check = 1

options = Options()
options.add_argument('--ignore-certificate-errors')  
options.add_argument('--disable-web-security')      

dele_type = "AMG"
input_acc = "14574"


driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)
driver.maximize_window()
error_messages = []
wait = WebDriverWait(driver, 10)  
wait2 = WebDriverWait(driver, 2)
try:

    #driver.get(f"http://cimes.seec.com.tw/{dele_type}/CimesDesktop.aspx")
    driver.get("http://172.16.16.156/ARG/CimesDesktop.aspx")
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
    element.click()
    driver.switch_to.frame("ifmMenu")
    element3 = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#navFRE.ctgr"))
    )
    actions = ActionChains(driver)
    actions.move_to_element(element3).click().perform()

    element4 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[righttype="WIPFRE"][cimesclass="EDC"]')))
    element4.click()


    button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="工程資料參數維護"]')))
    ActionChains(driver).move_to_element(button).click().perform()


    driver.switch_to.default_content()
    iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "win1")))
    driver.switch_to.frame(iframe)

    add_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnAdd")))
    actions2 = ActionChains(driver)
    actions2.move_to_element(add_button).click().perform()


    for item in data:
        time.sleep(1)
        dropdown_element2 = wait.until(EC.element_to_be_clickable((By.ID, "ddlDataType")))
        select_value2 = item["data_type"]
        script = f"""
            arguments[0].value = '{select_value2}';
            arguments[0].dispatchEvent(new Event('change'));
        """
        driver.execute_script(script, dropdown_element2)
        time.sleep(0.5)
        input_element6 = wait.until(EC.presence_of_element_located((By.ID, "ttbSamplesize")))
        driver.execute_script("arguments[0].value = arguments[1];", input_element6, item["get_quan"])

        input_element1 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#ttbParameter')))
        driver.execute_script("arguments[0].value = arguments[1];", input_element1, item["param_name"])
        
        input_element2 = wait.until(EC.presence_of_element_located((By.ID, "ttbDisplayName")))
        driver.execute_script("arguments[0].value = arguments[1];", input_element2, item["param_name2"])
        
        input_element3 = wait.until(EC.presence_of_element_located((By.ID, "ttbTarget")))
        driver.execute_script("arguments[0].value = arguments[1];", input_element3, item["fit"])
        
        input_element4 = wait.until(EC.presence_of_element_located((By.ID, "ttbUSL")))
        driver.execute_script("arguments[0].value = arguments[1];", input_element4, item["limit_high"])
        
        input_element5 = wait.until(EC.presence_of_element_located((By.ID, "ttbLSL")))
        driver.execute_script("arguments[0].value = arguments[1];", input_element5, item["limit_low"])
        


        rbt_enable = driver.find_element(By.ID, "rbtEnable")
        rbt_disable = driver.find_element(By.ID, "rbtDisable")

        if  item["state"] == "啟用":
            rbt_enable.click()  
        else:
            rbt_disable.click()  


        dropdown_element = wait.until(EC.element_to_be_clickable((By.ID, "ddlUnit")))
        select_value = item["unit"]
        script = f"""
            arguments[0].value = '{select_value}';
            arguments[0].dispatchEvent(new Event('change'));
        """
        driver.execute_script(script, dropdown_element)
        time.sleep(0.5)

        


        if select_value2 == "Number":
            dropdown_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ddlVariableType")))
            select = Select(dropdown_element3)
            select.select_by_index("1")

        time.sleep(1)

            
        save_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnSave")))
        actions3 = ActionChains(driver)
        actions3.move_to_element(save_button).click().perform()

        time.sleep(3)
        
        try:
            error_element = wait2.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".exceptionTitle")))
            error_text = error_element.text
            error_messages.append(error_text)

            error_info = {
                    "param_name":item["param_name"],
                    "data_type": item["data_type"],
                    "param_name2": item["param_name2"],
                    "fit": item["fit"],
                    "limit_high": item["limit_high"],
                    "limit_low": item["limit_low"],
                    "state": item["state"],
                    "unit": item["unit"],
                    "get_quan": item["get_quan"]
            }
            print(error_info)
            edit_data.append(error_info)
            print(edit_data)
            
            close_button = wait2.until(EC.element_to_be_clickable(
                (By.XPATH, "//td//input[@value='關閉' and @type='button']")
            ))
            driver.execute_script("arguments[0].click();", close_button)
            continue
        except:
            pass

        add_button2 = wait.until(EC.element_to_be_clickable((By.NAME, "btnAdd")))
        actions5 = ActionChains(driver)
        actions5.move_to_element(add_button2).click().perform()
#FOR始末點
    exit_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnExit")))
    actions4 = ActionChains(driver)
    actions4.move_to_element(exit_button).click().perform()
    time.sleep(1)

    print(edit_data)





    if check == 1:

        for item in edit_data:
            ttbParameter = wait.until(EC.element_to_be_clickable((By.NAME, "ttbParameter")))
            ttbParameter.clear()
            actions.move_to_element(ttbParameter).click().perform()
            ttbParameter.send_keys(item["param_name"])
            btnQuery = wait.until(EC.element_to_be_clickable((By.NAME, "btnQuery")))
            btnQuery.click()
            time.sleep(2)
            EditButton = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @class='CSGridEditButton']"))
            )
            actions.move_to_element(EditButton).click().perform()
            

            time.sleep(1)
            dropdown_element2 = wait.until(EC.element_to_be_clickable((By.ID, "ddlDataType")))
            select_value2 = item["data_type"]
            script = f"""
                arguments[0].value = '{select_value2}';
                arguments[0].dispatchEvent(new Event('change'));
            """
            driver.execute_script(script, dropdown_element2)
            time.sleep(0.5)
            input_element6 = wait.until(EC.presence_of_element_located((By.ID, "ttbSamplesize")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element6, item["get_quan"])


            
            input_element2 = wait.until(EC.presence_of_element_located((By.ID, "ttbDisplayName")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element2, item["param_name2"])
            
            input_element3 = wait.until(EC.presence_of_element_located((By.ID, "ttbTarget")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element3, item["fit"])
            
            input_element4 = wait.until(EC.presence_of_element_located((By.ID, "ttbUSL")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element4, item["limit_high"])
            
            input_element5 = wait.until(EC.presence_of_element_located((By.ID, "ttbLSL")))
            driver.execute_script("arguments[0].value = arguments[1];", input_element5, item["limit_low"])
            


            rbt_enable = driver.find_element(By.ID, "rbtEnable")
            rbt_disable = driver.find_element(By.ID, "rbtDisable")

            if  item["state"] == "啟用":
                rbt_enable.click()  
            else:
                rbt_disable.click()  


            dropdown_element = wait.until(EC.element_to_be_clickable((By.ID, "ddlUnit")))
            select_value = item["unit"]
            script = f"""
                arguments[0].value = '{select_value}';
                arguments[0].dispatchEvent(new Event('change'));
            """
            driver.execute_script(script, dropdown_element)
            time.sleep(0.5)

            


            if select_value2 == "Number":
                dropdown_element3 = wait.until(EC.element_to_be_clickable((By.ID, "ddlVariableType")))
                select = Select(dropdown_element3)
                select.select_by_index("1")

            time.sleep(1)

            
            save_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnSave")))
            actions3 = ActionChains(driver)
            actions3.move_to_element(save_button).click().perform()

            time.sleep(2)
            exit_button = wait.until(EC.element_to_be_clickable((By.NAME, "btnExit")))
            actions4 = ActionChains(driver)
            actions4.move_to_element(exit_button).click().perform()
            time.sleep(2)
    else:
        pass

    time.sleep(5)
finally:
    with open("error_messages.txt", "w", encoding="utf-8") as file:
        for message in error_messages:
            file.write(message + "\n")
    driver.quit()