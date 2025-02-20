from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from Utils.Excel import ExcelUtils
from Test_login.login import LoginAutomation
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class  PaymentModeAutomation:
    def __init__(self, driver):
        self.driver = driver
        
    def Payment_Mode(self):
        try:    
            function_name = "Payment Mode"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH,function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            sleep(8)
            self.driver.find_element(By.XPATH, '(//span[text()="Payment Mode"])').click()
            for row_num in range(2,valid_rows):
                Payment_Mode = sheet.cell(row=row_num, column=4).value
                print(Payment_Mode)
                short_code = sheet.cell(row=row_num, column=5).value
                print(short_code)
                Edit = sheet.cell(row=row_num, column=6).value
                print(Edit)
                Edit_data1 = sheet.cell(row=row_num, column=7).value
                print(Edit_data1)
                Edit_data2 = sheet.cell(row=row_num, column=8).value
                print(Edit_data2)
                Delete = sheet.cell(row=row_num, column=9).value
                print(Delete)
                self.driver.refresh()
                # Call add_Payment_Mode
                status=self.add_Payment_Mode(Payment_Mode,short_code)
                Desi_test_status,Desi_status,search_test_status,search_status,pay_id = status
                print(status)
                edit = self.edit_Payment_Mode(Edit,Edit_data1,Edit_data2,pay_id)
                Edit_test_status,Edit_status = edit   
                print(edit)
                sleep(3)
                delete = self.delete(Delete,pay_id)
                Delete_test_status,Delete_status = delete
                print(delete)
                test_status = "Fail" if "Fail" in [Desi_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
                Actual_status = ",".join([Desi_status, search_status, Edit_status, Delete_status])  
                sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                workbook.save(FILE_PATH)
            print("Payment Mode Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)         
        except Exception as e:
                print(f"Error during login: {e}") 
                        
    def add_Payment_Mode(self,Payment_Mode,short_code):
        sleep(2)
        self.driver.find_element(By.ID, 'add_paymode').click()   
        sleep(3)
        self.driver.find_element(By.ID, "mode_name").send_keys(Payment_Mode)
        self.driver.find_element(By.ID, "short_code").send_keys(short_code)
        sleep(5)
        self.driver.find_element(By.XPATH,'//button[@id="paymode_submit"]').click()
        sleep(10)
        url = LoginAutomation.url(self)
        base_url = url.split('php/')[0]+'php/pp/common/payment_mode/list'
        self.driver.get("https://php8.logimaxindia.com/titan_v7/admin/index.php/pp/common/payment_mode/list")
        try:
            # if self.driver.title == "Payment Mode - Payment_mode | TANEIRA":
                sleep(5)
                status=self.search(Payment_Mode)
                search_test_status,search_data,search_status,pay_id = status
                if Payment_Mode in search_data:
                    Desi_test_status = "Pass"
                    Desi_status = "Payment Mode Add successful"
                else:
                    Desi_test_status = "Fail"    
                    Desi_status = "Payment Mode Add Unsuccessful"   
            # else:       
            #     Desi_test_status = "Fail"    
            #     Desi_status = "Payment Mode Add Unsuccessful"
            #     search_test_status = "Fail" 
            #     search_status = "Data not added"
            #     pay_id = ' '
        # Handle exceptions              
        except Exception as e: 
            Desi_test_status = "Fail"
            Desi_status = f"Payment Mode Add Unsuccessful:{str(e)}"          
        status = Desi_test_status,Desi_status,search_test_status,search_status,pay_id
        return status
            
    def search(self,Payment_Mode):
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Payment_Mode)
        pay_id=' '
        try:
            for i in range(1, 6):
                cell_xpath = f'//table[@id="paymode_list"]//tbody//tr[{i}]//td[2]'
                data= self.driver.find_element(By.XPATH, cell_xpath)
                id=self.driver.find_element(By.XPATH,f'//table[@id="paymode_list"]//tbody//tr[{i}]//td[1]')
                pay_id = id.text
                search_data=data.text
                if (search_data==Payment_Mode):
                    search_test_status = "Pass"    
                    search_status=("search data successfull")
                    break
                else:
                    search_test_status = "Fail"
                    search_status=("search data Unsuccessfull") 
        except:
            search_data=[' ']
            search_test_status = "Fail"
            search_status=("search data Unsuccessfull")
        status = search_test_status,search_data,search_status,pay_id
        return status    
    def edit_Payment_Mode(self,Edit,Edit_data1,Edit_data2,pay_id):
        try:
        # Check if Dept_id is empty
            if pay_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit Payment Mode Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Edit), re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(pay_id)
                try:   
                    if  self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button'):
                        self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button').click()
                        sleep(3)
                        self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/ul/li[1]/a').click()
                        sleep(3)
                        self.driver.find_element(By.ID,"mode_name").clear()
                        self.driver.find_element(By.ID,"mode_name").send_keys(Edit_data1)
                        self.driver.find_element(By.ID,"short_code").clear()
                        self.driver.find_element(By.ID,"short_code").send_keys(Edit_data2)
                        self.driver.find_element(By.ID,"paymode_submit").click()
                        Edit_test_status = "Pass"
                        Edit_status = "Edit Payment Mode successful"
                except:
                    Edit_test_status = "Fail"
                    Edit_status = "Edit Payment Mode Unsuccessful"       
            elif re.match(r"No", str(Edit), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit Payment Mode Unsuccessful"   
        except:         
            Edit_test_status = "Fail"
            Edit_status = "Edit Payment Mode Unsuccessful"  
        edit =  Edit_test_status,Edit_status
        return edit  
            
    def delete(self,Delete,pay_id):
        try:
        # Check if Dept_id is empty
            if pay_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Delete Payment Mode Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Delete), re.IGNORECASE): 
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(pay_id) 
                try:
                    if self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button'):
                        self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/ul/li[2]/a').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH,'(//a[@onclick="deletePaymode(id,2)"])').click()
                        sleep(2)
                        Delete_test_status = "Pass"
                        Delete_status = "Delete Payment Mode  successful"
                except:
                    Delete_test_status = "Fail"
                    Delete_status = "Delete Payment Mode  Unsuccessful"         
            elif re.match(r"No", str(Delete), re.IGNORECASE):
                Delete_test_status = "Fail"
                Delete_status = "Delete Payment Mode  Unsuccessful" 
            delete = Delete_test_status,Delete_status
            return delete
        except: 
            Delete_test_status = "Fail"
            Delete_status = "Delete Payment Mode  Unsuccessful" 
        delete = Delete_test_status,Delete_status
        return delete
    