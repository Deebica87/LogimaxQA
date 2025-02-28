from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from Utils.Excel import ExcelUtils
from Test_Image.screenshot import image
import re
from openpyxl import load_workbook

FILE_PATH = ExcelUtils.file_path
class DepartmentAutomation:
    def __init__(self, driver):
        self.driver = driver
        
    def Department(self):
        try:    
            function_name = "Department"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH,function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
            sleep(5)
            self.driver.find_element(By.XPATH, '(//span[text()="Department"])').click()
            
            for row_num in range(2,valid_rows):
                Department = sheet.cell(row=row_num, column=4).value
                print(Department)
                Edit = sheet.cell(row=row_num, column=5).value
                print(Edit)
                Edit_data = sheet.cell(row=row_num, column=6).value
                print(Edit_data)
                Delete = sheet.cell(row=row_num, column=7).value
                print(Delete)
                self.driver.refresh()
                # Call add_department
                status=self.add_department(Department,function_name)
                Dept_test_status,Dept_status,search_test_status,search_status,Dept_id = status
                print(status)
                edit = self.edit_department(Edit_data,Edit,Dept_id,)
                Edit_test_status,Edit_status = edit   
                print(edit)
                sleep(3)
                delete = self.delete(Delete,Dept_id)
                Delete_test_status,Delete_status = delete
                print(delete)
                test_status = "Fail" if "Fail" in [Dept_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
                Actual_status = ",".join([Dept_status, search_status, Edit_status, Delete_status])  
                sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                workbook.save(FILE_PATH)
            print("Department Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)     
        except Exception as e:
                print(f"Error during login: {e}") 

    def add_department(self,Department,function_name):
        sleep(8)
        self.driver.find_element(By.ID, 'add_dpt').click()   
        sleep(3)
        self.driver.find_element(By.ID, "department").send_keys(Department)
        sleep(5)
        self.driver.find_element(By.XPATH, '(//a[@class="btn btn-success"])[1]').click()
        sleep(5)
        status=self.search(Department)
        search_test_status,search_data,search_status,Dept_id = status
        try:
            if Department in search_data:
                Dept_test_status = "Pass"
                Dept_status = "Department Add successful"
            else:
                path = image.screenshot(function_name)
                self.driver.get_screenshot_as_file(path)
                Dept_test_status = "Fail"    
                Dept_status = "Department Add Unsuccessful" 
                self.driver.get_screenshot_as_file('E:\\CRM\\Test\\Screenshots\\dept.png')    
        except Exception as e:  # Handle exceptions
                Dept_test_status = "Fail"
                Dept_status = f"Department Add Unsuccessful: {str(e)}"          
        status = Dept_test_status,Dept_status,search_test_status,search_status,Dept_id
        return status

    def search(self,Department):
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Department)
        Dept_id=' '
        try:
            for i in range(1, 6):
                cell_xpath = f'//table[@id="dept_list"]//tbody//tr[{i}]//td[2]'
                data= self.driver.find_element(By.XPATH, cell_xpath)
                id=self.driver.find_element(By.XPATH,f'//table[@id="dept_list"]//tbody//tr[{i}]//td[1]')
                Dept_id = id.text
                search_data=data.text
                if (search_data==Department):
                    break
            if (search_data==Department):
                search_test_status = "Pass"    
                search_status=("search data successfull")
            else:
                function_name='search'
                path = image.screenshot(function_name)
                self.driver.get_screenshot_as_file(path)
                search_test_status = "Fail"
                search_status=("search data Unsuccessfull") 
        except:
            search_data=[' ']
            search_test_status = "Fail"
            search_status=("search data Unsuccessfull")
        status = search_test_status,search_data,search_status,Dept_id
        return status    

    def edit_department(self,Edit_data,Edit,Dept_id):
        sleep(3)
        try:
        # Check if Dept_id is empty
            if Dept_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit department Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(Edit), re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Dept_id)
                try:   
                    if  self.driver.find_element(By.ID,"edit"):
                        self.driver.find_element(By.ID,"edit").click()
                        sleep(3)
                        self.driver.find_element(By.ID,"ed_dept").clear()
                        self.driver.find_element(By.ID,"ed_dept").send_keys(Edit_data)
                        self.driver.find_element(By.ID,"update_dept").click()
                        Edit_test_status = "Pass"
                        Edit_status = "Edit department successful"
                except:
                    function_name = 'Edit_Dept'
                    path = image.screenshot(function_name)
                    self.driver.get_screenshot_as_file(path)
                    Edit_test_status = "Fail"
                    Edit_status = "Edit department Unsuccessful"       
            elif re.match(r"No", str(Edit), re.IGNORECASE):
                Edit_test_status = "Pass"
                Edit_status = "Edit Option No in Excel sheet"  
        except:         
            Edit_test_status = "Fail"
            Edit_status = "Edit department Unsuccessful"  
        edit =  Edit_test_status,Edit_status
        return edit

    def delete(self,Delete,Dept_id):
        try:
            if Dept_id==' ':
                Delete_test_status = "Fail"
                Delete_status = "Delete department  Unsuccessful" 
                delete = Delete_test_status,Delete_status
                return delete     
            if  re.match(r"yes",str(Delete),re.IGNORECASE): 
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Dept_id) 
                try:
                    if self.driver.find_element(By.XPATH,'(//a[@class="btn btn-danger btn-del"])[1]'):
                        self.driver.find_element(By.XPATH,'(//a[@class="btn btn-danger btn-del"])[1]').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH,'(//a[@onclick="delete_dept(id,2)"])').click()
                        sleep(2)
                        Delete_test_status = "Pass"
                        Delete_status = "Delete department successful"
                except:
                    function_name = 'del_dept'
                    path = image.screenshot(function_name)
                    self.driver.get_screenshot_as_file(path)
                    Delete_test_status = "Fail"
                    Delete_status = "Delete department  Unsuccessful"         
            elif re.match(r"No", str(Delete), re.IGNORECASE):
                Delete_test_status = "Pass"
                Delete_status = "Delete Option No in Excel Sheet"
        except: 
            Delete_test_status = "Fail"
            Delete_status = "Delete department  Unsuccessful" 
        delete = Delete_test_status,Delete_status
        return delete
    
    
    def departmentname(self):
        function_name = "Department"
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]
        sleep(10)
        for row_num in range(2, valid_rows):
            # Define columns and dynamically fetch their values
            data = {
                "Department":4,
            }
            row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
            name = row_data["Department"]
            datas = name
            print(datas)
            return datas





