from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from Utils.Excel import ExcelUtils
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class ProfessionAutomation:
    def __init__(self, driver):
        self.driver = driver
        
    def Profession(self):
        try:    
            function_name = "Profession"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH,function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            sleep(8)
            self.driver.find_element(By.XPATH, '(//span[text()="Profession"])').click()

            for row_num in range(2,valid_rows):
                profession = sheet.cell(row=row_num, column=4).value
                print(profession)
                Edit = sheet.cell(row=row_num, column=5).value
                print(Edit)
                Edit_data = sheet.cell(row=row_num, column=6).value
                print(Edit_data)
                Delete = sheet.cell(row=row_num, column=7).value
                print(Delete)
                self.driver.refresh()
                # Call add_profession
                status=self.add_profession(profession)
                Desi_test_status,Desi_status,search_test_status,search_status,profe_id = status
                print(status)
                edit = self.edit_profession(Edit_data,Edit,profe_id)
                Edit_test_status,Edit_status = edit   
                print(edit)
                sleep(3)
                delete = self.delete(Delete,profe_id)
                Delete_test_status,Delete_status = delete
                print(delete)
                test_status = "Fail" if "Fail" in [Desi_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
                Actual_status = ",".join([Desi_status, search_status, Edit_status, Delete_status])  
                sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                workbook.save(FILE_PATH)
            print("profession Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)         
        except Exception as e:
                print(f"Error during login: {e}") 
                
    def add_profession(self,profession):
        sleep(2)
        self.driver.find_element(By.ID, 'add_professions').click()   
        sleep(3)
        self.driver.find_element(By.ID, "profession").send_keys(profession)
        sleep(5)
        self.driver.find_element(By.XPATH, '(//a[@class="btn btn-success"])[1]').click()
        sleep(5)
        status=self.search(profession)
        search_test_status,search_data,search_status,profe_id = status
        try:
            if profession in search_data:
                Desi_test_status = "Pass"
                Desi_status = "profession Add successful"
            else:
                Desi_test_status = "Fail"    
                Desi_status = "profession Add Unsuccessful"     
        except Exception as e:  # Handle exceptions
                Desi_test_status = "Fail"
                Desi_status = f"profession Add Unsuccessful: {str(e)}"          
        status = Desi_test_status,Desi_status,search_test_status,search_status,profe_id
        return status
            
    def search(self,profession):
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(profession)
        profe_id=' '
        try:
            for i in range(1, 6):
                cell_xpath = f'//table[@id="profession_list"]//tbody//tr[{i}]//td[2]'
                data= self.driver.find_element(By.XPATH, cell_xpath)
                id=self.driver.find_element(By.XPATH,f'//table[@id="profession_list"]//tbody//tr[{i}]//td[1]')
                profe_id = id.text
                search_data=data.text
                if (search_data==profession):
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
        status = search_test_status,search_data,search_status,profe_id
        return status    


    def edit_profession(self,Edit_data,Edit,profe_id):
        try:
        # Check if Dept_id is empty
            if profe_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit profession Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Edit), re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(profe_id)
                try:   
                    if  self.driver.find_element(By.ID,"edit"):
                        self.driver.find_element(By.ID,"edit").click()
                        sleep(3)
                        self.driver.find_element(By.ID,"ed_profession").clear()
                        self.driver.find_element(By.ID,"ed_profession").send_keys(Edit_data)
                        self.driver.find_element(By.ID,"update_profession").click()
                        Edit_test_status = "Pass"
                        Edit_status = "Edit profession successful"
                except:
                    Edit_test_status = "Fail"
                    Edit_status = "Edit profession Unsuccessful"       
            elif re.match(r"No", str(Edit), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit profession Unsuccessful"   
        except:         
            Edit_test_status = "Fail"
            Edit_status = "Edit profession Unsuccessful"  
        edit =  Edit_test_status,Edit_status
        return edit  
            
    def delete(self,Delete,profe_id):
        try:
        # Check if Dept_id is empty
            if profe_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Delete profession Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Delete), re.IGNORECASE): 
                try:
                    if self.driver.find_element(By.XPATH,'(//a[@class="btn btn-danger btn-del"])[1]'):
                        self.driver.find_element(By.XPATH,'(//a[@class="btn btn-danger btn-del"])[1]').click()
                        sleep(2)
                    self.driver.find_element(By.XPATH,'(//a[@onclick="delete_profession(id,2)"])').click()
                    sleep(2)
                    Delete_test_status = "Pass"
                    Delete_status = "Delete profession  successful"
                except:
                    Delete_test_status = "Fail"
                    Delete_status = "Delete profession  Unsuccessful"         
            elif re.match(r"No", str(Delete), re.IGNORECASE):
                Delete_test_status = "Fail"
                Delete_status = "Delete profession  Unsuccessful" 
            delete = Delete_test_status,Delete_status
            return delete
        except: 
            Delete_test_status = "Fail"
            Delete_status = "Delete profession  Unsuccessful" 
        delete = Delete_test_status,Delete_status
        return delete
    
    
    
               