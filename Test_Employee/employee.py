from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from time import sleep
from Utils.Excel import ExcelUtils
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class EmployeesAutomation:
    def __init__(self,driver):
        self.driver =driver

    def Employees(self):
        try:
            function_name = "Employees"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            # Navigating through the interface
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Employee').click()
            sleep(5)
            self.driver.find_element(By.XPATH, '(//span[text()="Employees"])').click()
            # Process each row
            for row_num in range(2, valid_rows):
                # Define columns and dynamically fetch their values
                data = {
                    "Profile Image": 4,
                    "First Name": 5,
                    "Last Name": 6,
                    "Date of Birth": 7,
                    "Employee Code": 8,
                    "Date of Joining": 9,
                    "Department": 10,
                    "Designation": 11,
                    "Address1": 12,
                    "Address2": 13,
                    "Address3": 14,
                    "Pincode": 15,
                    "Country": 16,
                    "State": 17,
                    "City": 18,
                    "Mobile": 19,
                    "Phone": 20,
                    "email": 21,
                    "User Name": 22,
                    "Password": 23,
                    "User Type": 24,
                    "Select Branch": 25,
                    "Edit":26,
                    "EditFirstName":27,
                    "EditDepartment":28,
                    "EditDesignation":29,
                    "EditMobile":30,
                    "Delete":31,
                }
                row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
                print(row_data)
                status=self.Employeeadd_data(row_data)
                try:
                    test_status,Actual_status = status
                    if test_status == 'Fail':
                        sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                        sheet.cell(row=row_num, column=3).value = Actual_status
                        workbook.save(FILE_PATH)
                except:    
                    Emp_test_status,Emp_status,search_test_status,search_status,Emptomer_id = status
                    print(status)  
                    edit = self.edit_Employees(row_data,Emptomer_id)
                    Edit_test_status,Edit_status = edit
                    print(edit)
                    delete =self.delete_Employees(row_data,Emptomer_id)
                    Delete_test_status,Delete_status = delete
                    test_status = "Fail" if "Fail" in [Emp_test_status,search_test_status,Edit_test_status,Delete_test_status] else "pass"
                    Actual_status = ",".join([Emp_status, search_status,Edit_status,Delete_status])  
                    print(test_status)
                    print(Actual_status)
                    sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                    sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                    workbook.save(FILE_PATH) 
            print("Classification Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)
        except Exception as e:
                print(f"Error during login: {e}") 
    
    def Employeeadd_data(self,row_data):
        self.driver.find_element(By.ID, "add_employee").click()
        Name = row_data["Profile Image"]
        Path = r"D:\CRM\Taneira\Image"  
        Image_path = f"{Path}\\{Name}.jpg"
        print(Image_path)
        self.driver.find_element(By.ID, "pp_emp_image").send_keys(Image_path)
        if row_data["First Name"] != None and row_data["Last Name"] != None:
            self.driver.find_element(By.ID, "firstname").send_keys(row_data["First Name"])
            self.driver.find_element(By.ID, "lastname").send_keys(row_data["Last Name"])
        else:  
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\First_name_fail.png") 
            return 'Fail','First Name & Last Name Should be Mandatory'
        if (row_data["Date of Birth"])!= None: 
            D_O_B = str(row_data["Date of Birth"])
            Date_of_birth = self.driver.find_element(By.ID, "date_of_birth")
            Date_of_birth.send_keys(D_O_B)
            Date_of_birth.send_keys(Keys.TAB)
        else:    
            return 'Fail','Date of Birth  is mandatory'
        sleep(5)
        self.driver.find_element(By.ID, "emp_code").send_keys(row_data["Employee Code"])
        sleep(3)
        D_O_J = str(row_data["Date of Joining"])
        Join_Date=self.driver.find_element(By.ID, "date_of_join")
        Join_Date.send_keys(D_O_B)
        Join_Date.send_keys(Keys.TAB)
        Join_Date.send_keys(Keys.TAB)
        if  row_data["Department"] != None:
            sleep(5)
            self.driver.find_element(By.XPATH, '//span[@id="select2-dept-container"]').click()
            state=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
            state.send_keys(row_data["Department"])
            state.send_keys(Keys.ENTER)  
        else:
            return 'Fail','Department should be mandatory field'
        sleep(5)
        if  row_data["Designation"] != None:
            print(row_data["Designation"])
            self.driver.find_element(By.XPATH, '//span[@id="select2-designation-container"]').click()
            state=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
            state.send_keys(row_data["Designation"])
            state.send_keys(Keys.ENTER)  
        else:
            return 'Fail','Designation should be mandatory field'
        sleep(4)
        if (row_data["Address1"])!=None:
            self.driver.find_element(By.ID, "address1").send_keys(row_data["Address1"])
        else:    
            return 'Fail','Address1 is mandatory' 
        self.driver.find_element(By.ID, "address2").send_keys(row_data["Address2"])
        self.driver.find_element(By.ID, "address3").send_keys(row_data["Address3"])
        sleep(2)
        self.driver.find_element(By.NAME, "pincode").send_keys(row_data["Pincode"])
        Country = row_data["Country"]
        State = row_data["State"]
        City = row_data["City"]
        print(Country)
        print(State)
        print(City)
        if row_data["Country"] != None and row_data["State"] != None and row_data["City"] != None:
            for i in [3, 3, 3]:  # Using a correct list for iteration
                clickable_xpath = f'(//span[@class="select2-selection__clear"])[{i}]'
                try:
                    # Locate the element
                    point = self.driver.find_element(By.XPATH, clickable_xpath)
                    # Ensure it's visible and clickable
                    self.driver.execute_script("arguments[0].scrollIntoView();", point)
                    sleep(1)
                    point.click()  
                    sleep(1)
                    actions = ActionChains(self.driver)
                    actions.send_keys(Keys.TAB).perform()
                    sleep(1)
                except Exception as e:
                    print(f"Error interacting with element {clickable_xpath}: {e}")
            sleep(5)
            self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[3]').click()
            sleep(3)
            print("1")
            country=self.driver.find_element(By.XPATH,"//input[@role='textbox']")
            country.send_keys(Country)
            country.send_keys(Keys.ARROW_DOWN)
            country.send_keys(Keys.RETURN)
            sleep(5)
            self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[4]').click()
            state=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
            state.send_keys(State)
            state.send_keys(Keys.ENTER)
            print(State)
            sleep(5)
            self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[5]').click()
            city =self.driver.find_element(By.XPATH,"//input[@role='textbox']")
            city.send_keys(City)
            city.send_keys(Keys.ENTER) 
        else:
            return "Fail","Country,State,City should be mandatory field"
        sleep(5)
        Mobile = self.driver.find_element(By.ID, "mobile")
        Mobile.send_keys(row_data["Mobile"])
        Mobile.send_keys(Keys.TAB)    
        sleep(5)
        self.driver.find_element(By.NAME, "phone").send_keys(row_data["Phone"]) 
        self.driver.find_element(By.ID, "email").send_keys(row_data["email"])
        self.driver.find_element(By.ID, "username").send_keys(row_data["User Name"])
        self.driver.find_element(By.ID, "passwd").send_keys(row_data["Password"])
        self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[6]').click()
        User_Type=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
        User_Type.send_keys(row_data["User Type"])
        User_Type.send_keys(Keys.ENTER)
        sleep(5)
        self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[7]').click()
        Branch=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
        Branch.send_keys(row_data["Select Branch"])
        Branch.send_keys(Keys.ENTER)
        print('3')
        self.driver.find_element(By.ID, 'emp_submit').click()
        print('4')
        status =self.search(row_data)
        search_test_status,search_data,search_status,Employee_id = status
        try:
            Employee=str(row_data["Mobile"])
            if  Employee in search_data:
                Emp_test_status = "Pass"
                Emp_status = "Employee Add successful"
            else:
                Emp_test_status = "Fail"    
                Emp_status = "Employee Add Unsuccessful"     
        except Exception as e:  # Handle exceptions
                Emp_test_status = "Fail"
                Emp_status = f"Employee Add Unsuccessful: {str(e)}"          
        status = Emp_test_status,Emp_status,search_test_status,search_status,Employee_id
        print(f'{status}3')
        return status
    
    def search(self,row_data):
        sleep(5)
        Employee=str(row_data["Mobile"])
        print(Employee)
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Employee)
        Employee_id=' '
        try:
            for i in range(1, 6):
                cell_xpath = f'//table[@id="emp_list"]//tbody//tr[{i}]//td[6]'
                data= self.driver.find_element(By.XPATH, cell_xpath)
                id=self.driver.find_element(By.XPATH,f'//table[@id="emp_list"]//tbody//tr[{i}]//td[1]')
                Employee_id = id.text
                search_data=data.text
                if (search_data==Employee):
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
        status = search_test_status,search_data,search_status,Employee_id
        print(status)
        return status 
    
    def edit_Employees(self,row_data,Emptomer_id):
        sleep(5)
        try:
            if Emptomer_id == None:
                Edit_test_status = "Fail"
                Edit_status = "Edit department Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit
            if  re.match(r"yes",str(row_data["Edit"]),re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Emptomer_id)
                try:
                    if self.driver.find_element(By.XPATH,'(//button[@class="btn btn-default dropdown-toggle"])[1]'):
                        self.driver.find_element(By.XPATH,'(//button[@class="btn btn-default dropdown-toggle"])[1]').click()
                        sleep(3)
                        self.driver.find_element(By.XPATH,'(//a[@class="btn-edit "])[1]').click()
                        self.driver.find_element(By.ID, "firstname").clear()
                        self.driver.find_element(By.ID, "firstname").send_keys(row_data["EditFirstName"])
                        self.driver.find_element(By.XPATH, '//span[@id="select2-dept-container"]').click()
                        state=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
                        state.send_keys(row_data["EditDepartment"])
                        state.send_keys(Keys.ENTER)
                        self.driver.find_element(By.XPATH, '//span[@id="select2-designation-container"]').click()
                        state=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
                        state.send_keys(row_data["EditDesignation"])
                        state.send_keys(Keys.ENTER)
                        sleep(5)
                        Mobile = self.driver.find_element(By.ID, "mobile")
                        Mobile.clear()
                        Mobile.send_keys(row_data["EditMobile"])
                        sleep(2)
                        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        sleep(20)
                        if self.driver.find_element(By.XPATH,'//button[@id="emp_submit"]'):
                           self.driver.find_element(By.XPATH,'//button[@id="emp_submit"]').click()
                           self.driver.find_element(By.XPATH,'//button[@id="emp_submit"]').click()
                        Edit_test_status = "Pass"
                        Edit_status = "Edit Customer successful"
                except:
                    Edit_test_status = "Fail"
                    Edit_status = "Edit Employees Unsuccessful"             
                        
            elif re.match(r"No", str(["Edit"]), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit Employees Unsuccessful"   
        except:         
            Edit_test_status = "Fail"
            Edit_status = "Edit Employees Unsuccessful"  
        edit =  Edit_test_status,Edit_status
        return edit  
    
    def delete_Employees(self,row_data,Emptomer_id):
        sleep(5)
        try:
        # Check if Dept_id is empty
            if Emptomer_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Delete Employees Unsuccessful"
                edit =  Edit_test_status,Edit_status
                return edit  # Return immediately
        # Check if Edit action is "yes" (case-insensitive)
            if re.match(r"yes", str(row_data['Delete']), re.IGNORECASE): 
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Emptomer_id) 
                try:
                    if self.driver.find_element(By.XPATH,'(//button[@class="btn btn-default dropdown-toggle"])[1]'):
                        self.driver.find_element(By.XPATH,'(//button[@class="btn btn-default dropdown-toggle"])[1]').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH,'(//a[@class="btn-del "])[1]').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH,'(//a[@onclick="deleteEmployee(delete_url = null,step=2)"])').click()
                        sleep(2)
                        Delete_test_status = "Pass"
                        Delete_status = "Delete Employees  successful"
                except:
                    Delete_test_status = "Fail"
                    Delete_status = "Delete Employees  Unsuccessful"         
            elif re.match(r"No", str(row_data['Delete']), re.IGNORECASE):
                Delete_test_status = "Fail"
                Delete_status = "Delete Employees  Unsuccessful" 
            delete = Delete_test_status,Delete_status
            return delete
        except: 
            Delete_test_status = "Fail"
            Delete_status = "Delete Employees  Unsuccessful" 
        delete = Delete_test_status,Delete_status
        return delete