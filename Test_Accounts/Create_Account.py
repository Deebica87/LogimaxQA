from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from time import sleep
from Utils.Excel import ExcelUtils
from Test_Customer.customer import CustomerAutomation
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import re


FILE_PATH = ExcelUtils.file_path
class CreateAccountAutomation:
    def __init__(self,driver):
        self.driver =driver

    def CreateAccount(self):
        try:
            function_name = "CreateAccount"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            # Navigating through the interface
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Manage Accounts').click()
            sleep(8)
            self.driver.find_element(By.XPATH, '(//span[text()="Create Account"])').click()
            # Process each row
            for row_num in range(2, valid_rows):
                data = {
                    "SelectBranch": 4,
                    "CustomerName": 5,
                    "PurchasePlan": 6,
                    "AccountName": 7,
                    "Category": 8,
                    "PayableAmount": 9,
                    "comment":10,
                    "UploadForm": 11,
                    "SelectReferral": 12,
                    "ReferralCode": 13,
                }
                row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
                print(row_data)
                # Url = LoginAutomation.url(self) 
                # Call add_Payment_Mode
                status=self.planAccount(row_data)
                
        except Exception as e:
                print(f"Error during login: {e}") 

    def planAccount(self,row_data):
        sleep(4)
        self.driver.find_element(By.XPATH,'//span[@id="select2-branch_select-container"]').click()
        branch = self.driver.find_element(By.XPATH,"(//input[@role='textbox'])[2]")
        branch.send_keys(row_data["SelectBranch"])
        branch.send_keys(Keys.ENTER)
        if row_data['CustomerName']!= None:
            Mobile_no = row_data['CustomerName']
        else:
            datas=CustomerAutomation.Customername(self)
            name,number = datas
            customer_name = f"{name}-{number}"
            first_four_digits = str(number)[:4]
        # Locate input field and type in number
        customer = self.driver.find_element(By.ID, 'mobile_number')
        customer.send_keys(number)  # Enter first part of number
        sleep(3)
        customer.send_keys(Keys.BACK_SPACE)  # Simulate user correction
        sleep(10)  # Allow dropdown to appear
        customer.send_keys(Keys.DOWN)
        customer.send_keys(Keys.ENTER)
        # # Fetch all elements from the dropdown
        # se_ver_list = self.driver.find_elements(By.XPATH, '//ul[@id="ui-id-1"]//li')
        # print(se_ver_list)
        # sleep(5)
        # for element in se_ver_list:
        #     print(f"Text: {element.text}")
        #     if element.text == customer_name:  # Replace this with your `customer_name` variable
        #         print(f"Clicking on: {element.text}")
        #         element.click()
        #         break
    
        # for element in se_ver_list:
        #     if element.text == customer_name:  # Replace this with your `customer_name` variable
        #         print(f"Clicking on: {element.text}")
        #         element.click()
        #         break
            
        sleep(5)
        self.driver.find_element(By.XPATH,'//span[@id="select2-scheme_select-container"]').click()
        plan = self.driver.find_element(By.XPATH,"(//input[@role='textbox'])[2]")
        plan.send_keys(row_data["PurchasePlan"])
        plan.send_keys(Keys.ENTER)
        sleep(3)
        self.driver.find_element(By.ID, 'account_name').send_keys(row_data['AccountName'])
        category=self.driver.find_element(By.XPATH,"(//input[@role='textbox'])")
        category.send_keys(row_data['Category'])
        category.send_keys(Keys.ENTER)
        sleep(2)
        Pay = self.driver.find_element(By.ID, "chosen_payable")
        Pay.send_keys(row_data['PayableAmount'])
        Pay.send_keys(Keys.TAB)
        sleep(2)
        self.driver.find_element(By.XPATH, '(//button[@class="btn-close"])[2]').click()
        self.driver.find_element(By.ID, 'remark_open').send_keys(row_data['comment'])
        UploadForm = row_data["UploadForm"]
        Path = r"D:\CRM\Taneira\Image"  
        Image_path = f"{Path}\\{UploadForm}.jpg"
        print(Image_path)
        self.driver.find_element(By.XPATH,'//input[@name="pp_det_image"]').send_keys(Image_path)
        self.driver.find_element(By.XPATH, '(//button[@type="button"])[1]').click()
        sleep(10)
        self.driver.find_element(By.ID, 'acc_submit').click()
        sleep(10)
        data = self.Payment(row_data,number)
        
    def Payment(self,row_data,number):
        self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
        sleep(5)
        self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Payment').click()
        sleep(8)
        self.driver.find_element(By.XPATH, '(//span[text()="Payments"])').click()
        sleep(5)
        # Process each row
        self.driver.find_element(By.ID, 'add_post_payment').click()
        # Locate input field and type in number
        customer = self.driver.find_element(By.ID, 'mobile_number')
        customer.send_keys(number)  # Enter first part of number
        sleep(3)
        customer.send_keys(Keys.BACK_SPACE)  # Simulate user correction
        sleep(10)  # Allow dropdown to appear
        customer.send_keys(Keys.DOWN)
        customer.send_keys(Keys.ENTER)
        self.driver.find_element(By.ID, 'select2-scheme_account-container')
        select_Acc_No=self.driver.find_element(By.XPATH,"(//input[@role='textbox'])")
        customer.send_keys(Keys.DOWN)
        select_Acc_No.send_keys(Keys.RETURN)
        self.driver.find_element(By.ID, 'select2-select_branch-container')
        Branch=self.driver.find_element(By.XPATH,"(//input[@role='textbox'])")
        Branch.send_keys(Keys.row_data[''])
        Branch.send_keys(Keys.RETURN)
        self.driver.find_element(By.XPATH,'make_pay_cash').send_keys(row_data[''])
        self.driver.find_element(By.ID, 'pay_save').click()
        