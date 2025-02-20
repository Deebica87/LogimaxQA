from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from Utils.Excel import ExcelUtils
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class DesignationAutomation:
    def __init__(self, driver):
        self.driver = driver
        
    def Designation(self):
        try:    
            function_name = "Designation"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH,function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
            sleep(5)
            self.driver.find_element(By.XPATH, '(//span[text()="Designation"])').click()

            for row_num in range(2,valid_rows):
                Designation = sheet.cell(row=row_num, column=4).value
                print(Designation)
                Edit = sheet.cell(row=row_num, column=5).value
                print(Edit)
                Edit_data = sheet.cell(row=row_num, column=6).value
                print(Edit_data)
                Delete = sheet.cell(row=row_num, column=7).value
                print(Delete)
                self.driver.refresh()
                # Call add_Designation
                status=self.add_designation(Designation)
                Desi_test_status,Desi_status,search_test_status,search_status,Desi_id = status
                print(status)
                edit = self.edit_designation(Edit_data,Edit,Desi_id)
                Edit_test_status,Edit_status = edit   
                print(edit)
                sleep(3)
                delete = self.delete(Delete,Desi_id)
                Delete_test_status,Delete_status = delete
                print(delete)
                test_status = "Fail" if "Fail" in [Desi_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
                Actual_status = ",".join([Desi_status, search_status, Edit_status, Delete_status])  
                sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                workbook.save(FILE_PATH)
            print("Designation Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)         
        except Exception as e:
                print(f"Error during login: {e}") 
            
                
    def add_designation(self,Designation):
        sleep(2)
        self.driver.find_element(By.ID, 'add_designation').click()   
        sleep(3)
        self.driver.find_element(By.ID, "designation").send_keys(Designation)
        sleep(5)
        self.driver.find_element(By.XPATH, '(//a[@class="btn btn-success"])[1]').click()
        sleep(5)
        status=self.search(Designation)
        search_test_status,search_data,search_status,Desi_id = status
        try:
            if Designation in search_data:
                Desi_test_status = "Pass"
                Desi_status = "Designation Add successful"
            else:
                Desi_test_status = "Fail"    
                Desi_status = "Designation Add Unsuccessful"     
        except Exception as e:  # Handle exceptions
                Desi_test_status = "Fail"
                Desi_status = f"Designation Add Unsuccessful: {str(e)}"          
        status = Desi_test_status,Desi_status,search_test_status,search_status,Desi_id
        return status
            
    def search(self,Designation):
            self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Designation)
            Desi_id=' '
            try:
                for i in range(1, 6):
                    cell_xpath = f'//table[@id="design_list"]//tbody//tr[{i}]//td[2]'
                    data= self.driver.find_element(By.XPATH, cell_xpath)
                    id=self.driver.find_element(By.XPATH,f'//table[@id="design_list"]//tbody//tr[{i}]//td[1]')
                    Desi_id = id.text
                    search_data=data.text
                    if(search_data==Designation):
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
            status = search_test_status,search_data,search_status,Desi_id
            return status    

    def edit_designation(self,Edit_data,Edit,Desi_id):
        try:
        # Check if Dept_id is empty
            if Desi_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit designation Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Edit), re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Desi_id)
                try:   
                    if  self.driver.find_element(By.ID,"edit"):
                        self.driver.find_element(By.ID,"edit").click()
                        sleep(3)
                        self.driver.find_element(By.ID,"ed_design").clear()
                        self.driver.find_element(By.ID,"ed_design").send_keys(Edit_data)
                        self.driver.find_element(By.ID,"update_design").click()
                        Edit_test_status = "Pass"
                        Edit_status = "Edit designation successful"
                except:
                    Edit_test_status = "Fail"
                    Edit_status = "Edit designation Unsuccessful"       
            elif re.match(r"No", str(Edit), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit designation Unsuccessful"   
        except:         
            Edit_test_status = "Fail"
            Edit_status = "Edit designation Unsuccessful"  
        edit =  Edit_test_status,Edit_status
        return edit        
            

    def delete(self,Delete,Desi_id):
        try:
        # Check if Dept_id is empty
            if Desi_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit department Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Delete), re.IGNORECASE): 
                try:
                    if self.driver.find_element(By.XPATH,'(//a[@class="btn btn-danger btn-del"])[1]'):
                        self.driver.find_element(By.XPATH,'(//a[@class="btn btn-danger btn-del"])[1]').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH,'(//a[@onclick="delete_design(id,2)"])').click()
                        sleep(2)
                        Delete_test_status = "Pass"
                        Delete_status = "Delete designation  successful"
                except:
                    Delete_test_status = "Fail"
                    Delete_status = "Delete designation  Unsuccessful"         
            elif re.match(r"No", str(Delete), re.IGNORECASE):
                Delete_test_status = "Fail"
                Delete_status = "Delete designation  Unsuccessful" 
            delete = Delete_test_status,Delete_status
            return delete
        except: 
            Delete_test_status = "Fail"
            Delete_status = "Delete designation  Unsuccessful" 
        delete = Delete_test_status,Delete_status
        return delete
    
    