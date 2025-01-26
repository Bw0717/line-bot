import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QStackedWidget, QTabWidget, QRadioButton, QGroupBox, QCheckBox, QFileDialog
from PyQt5.QtCore import Qt
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
from utils.pd import read_excel_and_group_by_hierarchy
from utils.pd import read_excel_and_auto_work_order
import keyboard
def auto_work_order(file_path,dele_type):
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
    data_inputs = read_excel_and_auto_work_order(file_path)

    try:
        driver.get(f"http://cimes.seec.com.tw/{dele_type}/Security/CimesUserLogin.aspx")
        

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


        keyboard.press_and_release('ctrl+-')
        
        for data_input in data_inputs:
            input_element3 = wait.until(EC.presence_of_element_located((By.ID, "CimesInputBox")))
            input_element3.clear()
            time.sleep(1)
            input_element3.send_keys(data_input)
            input_element3.send_keys(Keys.RETURN)
            time.sleep(1.5)
            enable_input = wait.until(EC.presence_of_element_located((By.ID, "ttbUnCreateQty")))
            value = enable_input.get_attribute('value')
            if value == '0':
                continue
            else:
                pass
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

def automation_input(file_path,dele_type,check):
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
        username_input = WebDriverWait(driver, 10).until(
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
        with open("error_messages.txt", "w", encoding="utf-8") as file:
            for message in error_messages:
                file.write(message + "\n")
        time.sleep(5)
    finally:

        driver.quit()


def automation_parameter(file_path,check,dele_type):


    data = read_excel_and_process(file_path)
    edit_data = []




    options = Options()
    options.add_argument('--ignore-certificate-errors')  
    options.add_argument('--disable-web-security')      


    input_acc = "14574"


    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    error_messages = []
    wait = WebDriverWait(driver, 10)  
    wait2 = WebDriverWait(driver, 2)
    try:

        driver.get(f"http://cimes.seec.com.tw/{dele_type}/CimesDesktop.aspx")
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



class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # 設置視窗標題和大小
        self.setWindowTitle("MES批量自動化")
        self.resize(600, 400)
        global check
        check = 0
        # 創建 QStackedWidget 來管理多個界面
        self.stacked_widget = QStackedWidget(self)

        # 創建主界面和參數資料界面
        self.main_page = self.create_main_page()
        self.param_page = self.create_param_page()

        # 添加界面到 QStackedWidget 中
        self.stacked_widget.addWidget(self.main_page)
        self.stacked_widget.addWidget(self.param_page)

        # 設置 QStackedWidget 作為主視窗的布局
        layout = QVBoxLayout()
        layout.addWidget(self.stacked_widget)
        self.setLayout(layout)

        footer_label = QLabel("Version: 1.0.0 | Writer: CHIH-HSIANG", self)
        footer_label.setAlignment(Qt.AlignRight)  # 右對齊
        layout.addWidget(footer_label, alignment=Qt.AlignBottom | Qt.AlignRight)  # 置底並右對齊

    def create_main_page(self):
        """主頁面：包含兩個按鈕"""
        main_page = QWidget()
        
        # 主頁標籤
        label = QLabel("歡迎使用批量自動化工具", main_page)
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 20px; font-weight: bold; color: #333;")

        # 按鈕：參數資料
        button1 = QPushButton("參數資料", main_page)
        button1.setFixedSize(200, 100)
        button1.setStyleSheet("font-size: 16px; background-color: #4CAF50; color: white; border-radius: 10px;")
        button1.clicked.connect(self.on_button1_click)

        # 按鈕：開立工單
        button2 = QPushButton("結束", main_page)
        button2.setFixedSize(200, 100)
        button2.setStyleSheet("font-size: 16px; background-color: #FF5722; color: white; border-radius: 10px;")
        button2.clicked.connect(self.on_button2_click)

        # 布局
        layout = QVBoxLayout()
        layout.addWidget(label)
        
        # 按鈕區域
        button_layout = QHBoxLayout()
        button_layout.addWidget(button1)
        button_layout.addWidget(button2)
        button_layout.setSpacing(20)
        layout.addLayout(button_layout)

        # 設置主頁面布局
        main_page.setLayout(layout)

        return main_page
    def on_create_work_order(self):
        """處理建立工單按鈕點擊事件"""
        # 這裡是你需要在建立工單按鈕點擊時執行的邏輯
        print("開始建立工單...")
    def create_param_page(self):
        """參數資料頁面：點擊按鈕進入的界面"""
        param_page = QWidget()

        # 創建 Tab 控制的界面
        tab_widget = QTabWidget(param_page)

        # 創建 Tab1頁面（建立參數）
        tab1 = QWidget()
        tab1_layout = QVBoxLayout()
        label1 = QLabel("請選擇一個選項:", tab1)
        group_box = QGroupBox("選擇單位", tab1)
        group_box2 = QGroupBox("選擇要執行的服務", tab1)
        group_box2.resize(180, 130)
        group_box2.move(300, 110)

        self.radio_arg = QRadioButton("ARG", group_box)
        self.radio_amg = QRadioButton("AMG", group_box)
        self.radio_am2 = QRadioButton("AM2", group_box)

        self.radio_P = QRadioButton("建立參數", group_box2)
        self.radio_I = QRadioButton("匯入參數", group_box2)

        # 設置默認選中
        self.radio_arg.setChecked(False)
        self.radio_P.setChecked(False)
        self.radio_arg.toggled.connect(self.on_radio_button_changed)
        self.radio_amg.toggled.connect(self.on_radio_button_changed)
        self.radio_am2.toggled.connect(self.on_radio_button_changed)
        self.radio_P.toggled.connect(self.on_radio_button_changed2)
        self.radio_I.toggled.connect(self.on_radio_button_changed2)

        # 設置圓形選項按鈕的排列
        radio_layout = QHBoxLayout()
        radio_layout.addWidget(self.radio_arg)
        radio_layout.addWidget(self.radio_amg)
        radio_layout.addWidget(self.radio_am2)
        group_box.setLayout(radio_layout)

        radio_layout2 = QVBoxLayout()
        radio_layout2.addWidget(self.radio_P)
        radio_layout2.addWidget(self.radio_I)
        group_box2.setLayout(radio_layout2)

        self.selected_label = QLabel("目前選擇: ", tab1)
        self.checkbox = QCheckBox("覆蓋已建立參數", tab1)
        self.checkbox.setChecked(False)
        self.checkbox.toggled.connect(self.on_checkbox_toggled)

        # 選擇檔案路徑按鈕
        file_button = QPushButton("選擇檔案路徑", tab1)
        file_button.setFixedSize(200, 50)
        file_button.setStyleSheet("font-size: 14px; background-color: #FF5722; color: white; border-radius: 10px;")
        file_button.clicked.connect(self.on_select_file)

        # Tab1的布局
        tab1_layout.addWidget(label1)
        tab1_layout.addWidget(group_box)
        tab1_layout.addWidget(self.selected_label)
        tab1_layout.addWidget(self.checkbox)
        tab1_layout.addWidget(file_button)
        tab1.setLayout(tab1_layout)

        # 創建 Tab2頁面（建立工單）
        tab2 = QWidget()
        tab2_layout = QVBoxLayout()
        label2 = QLabel("建立工單", tab2)

        # 創建圓形按鈕
        group_box2 = QGroupBox("選擇單位", tab2)
        self.radio_arg2 = QRadioButton("ARG", group_box2)
        self.radio_amg2 = QRadioButton("AMG", group_box2)
        self.radio_am22 = QRadioButton("AM2", group_box2)

        # 設置默認選中
        self.radio_arg2.setChecked(False)

        # 設置圓形選項按鈕的排列
        radio_layout2 = QHBoxLayout()
        radio_layout2.addWidget(self.radio_arg2)
        radio_layout2.addWidget(self.radio_amg2)
        radio_layout2.addWidget(self.radio_am22)
        group_box2.setLayout(radio_layout2)

        self.selected_label2 = QLabel("目前選擇:", tab2)
        self.radio_arg2.toggled.connect(self.on_radio_button_changed3)
        self.radio_amg2.toggled.connect(self.on_radio_button_changed3)
        self.radio_am22.toggled.connect(self.on_radio_button_changed3)

        # 選擇檔案路徑按鈕
        file_button2 = QPushButton("選擇檔案路徑", tab2)
        file_button2.setFixedSize(200, 50)
        file_button2.setStyleSheet("font-size: 14px; background-color: #FF5722; color: white; border-radius: 10px;")
        file_button2.clicked.connect(self.on_select_file2)

        # Tab2的布局
        tab2_layout.addWidget(label2)
        tab2_layout.addWidget(group_box2)
        tab2_layout.addWidget(file_button2)
        tab2.setLayout(tab2_layout)

        # 添加 Tab1 和 Tab2 頁面
        tab_widget.addTab(tab1, "建立參數")
        tab_widget.addTab(tab2, "建立工單")  # 新增的 Tab2

        # 設置按鈕區域
        button_layout = QHBoxLayout()

        start_button = QPushButton("開始執行", param_page)
        start_button.setFixedSize(180, 70)
        start_button.setStyleSheet("font-size: 16px; background-color: #4CAF50; color: white; border-radius: 10px;")
        start_button.clicked.connect(self.on_start_button_click2)

        back_button = QPushButton("返回主頁", param_page)
        back_button.setFixedSize(180, 70)
        back_button.setStyleSheet("font-size: 16px; background-color: #f44336; color: white; border-radius: 10px;")
        back_button.clicked.connect(self.on_back_button_click)

        button_layout.addWidget(start_button)
        button_layout.addWidget(back_button)
        button_layout.setSpacing(20)

        layout = QVBoxLayout()
        layout.addWidget(tab_widget)
        layout.addLayout(button_layout)
        param_page.setLayout(layout)

        return param_page



    def on_button1_click(self):
        """切換到參數資料頁面"""
        self.stacked_widget.setCurrentIndex(1)

    def on_button2_click(self):
        """開立工單功能"""
        QApplication.quit()
    def on_radio_button_changed2(self):
        global select_type
        select_type = None
        if self.radio_P.isChecked():
            select_type = "建立參數"
        elif self.radio_I.isChecked():
            select_type = "匯入參數"

    def on_radio_button_changed(self):
        global dele_type
        dele_type = None
        """處理圓形選項按鈕的選擇改變"""
        if self.radio_arg.isChecked():
            self.selected_label.setText("目前選擇: ARG")
            dele_type = "ARG"
        elif self.radio_amg.isChecked():
            self.selected_label.setText("目前選擇: AMG")
            dele_type = "AMG"
        elif self.radio_am2.isChecked():
            self.selected_label.setText("目前選擇: AM2")
            dele_type = "AM2"


    def on_checkbox_toggled(self):
        global check
        check = None
        """勾選框狀態改變"""
        if self.checkbox.isChecked():
            print("已選擇覆蓋已建立參數")
            check = 1
        else:
            print("未選擇覆蓋已建立參數")
            check = 0
    def on_select_file(self):
        """選擇檔案功能"""
        global file_path
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "選擇檔案", "", "All Files (*.*)")
        if file_path:
            print(f"選擇的檔案路徑是: {file_path}")
        
    def on_select_file2(self):
        """選擇檔案功能"""
        global file_path2
        file_dialog = QFileDialog(self)
        file_path2, _ = file_dialog.getOpenFileName(self, "選擇檔案", "", "All Files (*.*)")
        if file_path2:
            print(f"選擇的檔案路徑是: {file_path2}")

    def on_start_button_click2(self):
        """開始執行按鈕"""
        print("開始執行")
        try:
            print(file_path2)
            auto_work_order(file_path=file_path2,dele_type=dele_type)
        except:
            try:
                print(select_type)
                if select_type == "建立參數":
                    try:
                        automation_parameter(file_path = file_path,check = check,dele_type = dele_type)
                    except:
                        print("系統異常")
                elif select_type == "匯入參數":
                    try:
                        automation_input(file_path = file_path,dele_type = dele_type,check = check)
                    except:
                        print("系統異常")
            except:
                pass


    def on_back_button_click(self):
        """返回主頁按鈕"""
        self.stacked_widget.setCurrentIndex(0)

    def on_radio_button_changed3(self):
        """圓形選項按鈕狀態改變時更新顯示"""
        # 檢查哪個圓形選項按鈕被選中，並更新顯示的文本
        global dele_type
        if self.radio_arg2.isChecked():
            self.selected_label2.setText("目前選擇: ARG")
            dele_type = "ARG"
        elif self.radio_amg2.isChecked():
            self.selected_label2.setText("目前選擇: AMG")
            dele_type = "AMG"
        elif self.radio_am22.isChecked():
            self.selected_label2.setText("目前選擇: AM2")
            dele_type = "AM2"
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
