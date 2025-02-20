from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from Utils.Excel import ExcelUtils
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class ClassificationAutomation:
    def __init__(self, driver):
        self.driver = driver

    def classification(self):
        try:
            function_name = "classification"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            
            # Navigating through the interface
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
            
            sleep(8)
            self.driver.find_element(By.XPATH, '(//span[text()="Scheme Classification"])').click()

            # Process each row
            for row_num in range(2, valid_rows):
                # Define columns and dynamically fetch their values
                data = {
                    "Classification_Name": 4,
                    "image":5,
                    "Edit":6,
                    "Edit_data1":7,
                    "Edit_data2":8,
                    "Delete":9,
                }
                row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
            
                # Call add_Payment_Mode
                status=self.add_Classification(row_data)
                Desi_test_status, Desi_status, search_test_status, search_status, Classification_id = status
                print(status)
                edit = self.edit_classification(row_data,Classification_id)
                Edit_test_status,Edit_status = edit 
                print(edit)
                Delete = self.Delete_classification(row_data,Classification_id)
                Delete_test_status,Delete_status = Delete
                print(Delete)
                test_status = "Fail" if "Fail" in [Desi_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
                Actual_status = ",".join([Desi_status, search_status, Edit_status, Delete_status])  
                sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                workbook.save(FILE_PATH)
            print("Classification Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)    
        except Exception as e:
                print(f"Error during login: {e}")   
    
    def add_Classification(self,row_data):
        try:
            sleep(2)
            classification=row_data["Classification_Name"]
            image = row_data["image"]
            # Click 'Add classification' button
            self.driver.find_element(By.ID, 'add_cls').click()
            sleep(3)
            self.driver.find_element(By.ID, 'clsfy').send_keys(classification)
            self.driver.find_element(By.ID, 'sch_clsfy_img').send_keys(image)
            self.driver.find_element(By.ID, 'add_clsfy').click()
            sleep(2)
            # Search for the newly added classification
            search_status = self.search(row_data)
            search_test_status, search_data, search_msg, class_id = search_status
            # Validate search results
            if classification in search_data:
                return "Pass", "classification Add successful", search_test_status, search_msg, class_id
            else:
                return "Fail", "classification Add unsuccessful", search_test_status, search_msg, class_id

        except Exception as e:
            # Handle exceptions and return failure statuses
            return (
                "Fail",
                f"classification Add unsuccessful: {str(e)}",
                "Fail",
                "Search failed due to an exception",
                "",
            )
        
    def search(self, row_data):
        try:
            # Directly access 'classification_Name' from the dictionary
            classification=row_data["Classification_Name"]
        except KeyError:
            # Handle case where the key doesn't exist
            classification = ""
        
        # Enter the classification name into the search field
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(classification)
        
        class_id = ' '
        search_test_status, search_status, search_data = "Fail", "search data Unsuccessful", [' ']
        
        try:
            # Iterate through table rows to find the matching classification name
            for i in range(1, 6):
                row = self.driver.find_element(By.XPATH, f'//table[@id="sch_clsfy_list"]//tbody//tr[{i}]')
                search_data = row.find_element(By.XPATH, './td[2]').text
                class_id = row.find_element(By.XPATH, './td[1]').text
                if search_data == classification:
                    search_test_status, search_status = "Pass", "search data successful"
                    break
                else:
                    search_test_status, search_status = "Fail", "search data Unsuccessful"
        except:
            search_data=[' ']
            search_test_status = "Fail"
            search_status=("search data Unsuccessfull")
        
        return search_test_status, search_data, search_status, class_id
    
    def edit_classification(self, row_data, class_id):
        try:
            # Check if class_id is empty (not found)
            if class_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit classification Unsuccessful: classification ID not found"
                return Edit_test_status, Edit_status  # Return immediately

            # Simplified Edit check
            Edit = row_data["Edit"]

            # Check if Edit action is "yes"
            if re.match(r"yes", str(Edit), re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(class_id)
                try:
                    # Check if the edit button exists for the row
                    edit_button = self.driver.find_element(By.ID, 'edit')
                    if edit_button.is_displayed():
                        edit_button.click()
                        sleep(3)
                        # Clear existing data and send new data from row_data for editing
                        self.driver.find_element(By.ID, "ed_clsfy").clear()
                        self.driver.find_element(By.ID, "ed_clsfy").send_keys(row_data["Edit_data1"])
                        self.driver.find_element(By.ID, "edit_sch_clsfy_img").send_keys(row_data["Edit_data2"])
                        # Submit the form after editing
                        self.driver.find_element(By.ID, "update_clsfy").click()
                        Edit_test_status = "Pass"
                        Edit_status = "Edit classification successful"
                    else:
                        Edit_test_status = "Fail"
                        Edit_status = "Edit button not found"
                except Exception as e:
                    Edit_test_status = "Fail"
                    Edit_status = f"Edit classification Unsuccessful: {str(e)}"

            # If Edit action is "No"
            elif re.match(r"No", str(Edit), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit classification Unsuccessful: Edit action is 'No'"

            else:
                Edit_test_status = "Fail"
                Edit_status = "Edit classification Unsuccessful: Invalid Edit action"
            
        except Exception as e:
            Edit_test_status = "Fail"
            Edit_status = f"Edit classification Unsuccessful: {str(e)}"
        
        return Edit_test_status, Edit_status
    
    def Delete_classification(self,row_data,bank_id):
        try:
            # Check if bank_id is empty (not found)
            if bank_id == ' ':
                Delete_test_status = "Fail"
                Delete_status = "Delete Bank Unsuccessful: Bank ID not found"
                return Delete_test_status, Delete_status  # Return immediately
            # Simplified Delete check
            Delete = row_data["Delete"]
            # Check if Delete action is "yes"
            if re.match(r"yes", str(Delete), re.IGNORECASE):
                # Search for the bank_id in the list
                sleep(4)
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(bank_id)
                try:
                    # Check if the delete button exists for the row and click it
                    delete_button = self.driver.find_element(By.XPATH, '(//a[@class="btn btn-danger btn-del"])[1]')
                    if delete_button.is_displayed():
                        delete_button.click()
                        sleep(2)
                        # Confirm deletion
                        self.driver.find_element(By.XPATH, '(//a[@onclick="delete_classification_scheme(id,2)"])').click()
                        sleep(2)
                        Delete_test_status = "Pass"
                        Delete_status = "Delete Bank successful"
                    else:
                        Delete_test_status = "Fail"
                        Delete_status = "Delete button not found"
                except Exception as e:
                    Delete_test_status = "Fail"
                    Delete_status = f"Delete Bank Unsuccessful: {str(e)}"
                    
            elif re.match(r"No", str(Delete), re.IGNORECASE):
                Delete_test_status = "Fail"
                Delete_status = "Delete Bank Unsuccessful: Delete action is 'No'"

            else:
                Delete_test_status = "Fail"
                Delete_status = "Delete Bank Unsuccessful: Invalid Delete action"
            
        except Exception as e:
            Delete_test_status = "Fail"
            Delete_status = f"Delete Bank Unsuccessful: {str(e)}"
        
        return Delete_test_status, Delete_status
