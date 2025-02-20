from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from Utils.Excel import ExcelUtils
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class BankAutomation:
    def __init__(self, driver):
        self.driver = driver

    def bank(self):
        try:
            function_name = "Bank"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            # Navigating through the interface
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
            sleep(8)
            self.driver.find_element(By.XPATH, '(//span[text()="Bank"])').click()

            # Process each row
            for row_num in range(2, valid_rows):
                # Define columns and dynamically fetch their values
                data = {
                    "Bank_Name": 4,
                    "Short_Code": 5,
                    "Acc_NO": 6,
                    "IFSC": 7,
                    "Bank_Address": 8,
                    "Edit": 9,
                    "Edit_Data1": 10,
                    "Edit_Data2": 11,
                    "Edit_Data3": 12,
                    "Edit_Data4": 13,
                    "Edit_Data5": 14,
                    "Delete": 15,
                }
                row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}

                # Call add_Payment_Mode
                status=self.add_bank(row_data)
                Desi_test_status, Desi_status, search_test_status, search_status, bank_id = status
                print(status)
                edit = self.edit_bank(row_data,bank_id)
                Edit_test_status,Edit_status = edit 
                print(edit)
                Delete = self.Delete_bank(row_data,bank_id)
                Delete_test_status,Delete_status = Delete
                print(Delete)
                test_status = "Fail" if "Fail" in [Desi_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
                Actual_status = ",".join([Desi_status, search_status, Edit_status, Delete_status])  
                sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                workbook.save(FILE_PATH)
            print("Bank Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)    
        except Exception as e:
                print(f"Error during login: {e}")         
                
    def add_bank(self,row_data):
        try:
            sleep(2)
            # Click 'Add Bank' button
            self.driver.find_element(By.ID, 'add_bnk').click()
            sleep(3)
             # Map the required fields to their respective IDs
            field_mappings = {
                "Bank_Name": "bank_name",
                "Short_Code": "short_code",
                "Acc_NO": "acc_number",
                "IFSC": "ifsc_code",
                "Bank_Address": "address",
            }
            # Populate form fields dynamically
            for key, field_id in field_mappings.items():
                if key in row_data and row_data[key]:  # Check if the key exists and is not None
                    print(row_data[key])
                    self.driver.find_element(By.ID, field_id).send_keys(row_data[key])
             # Submit the form
            sleep(5)
            self.driver.find_element(By.XPATH, '(//button[@class="btn btn-primary"])[1]').click()

            # Navigate to Bank list
            payment_mode_url = "https://php8.logimaxindia.com/titan_v7/admin/index.php/pp/common/bank/list"
            self.driver.get(payment_mode_url)
            sleep(5)

            # Search for the newly added Bank
            search_status = self.search(row_data)
            search_test_status, search_data, search_msg, bank_id = search_status

            # Validate search results
            if row_data["Bank_Name"] in search_data:
                return "Pass", "Bank Add successful", search_test_status, search_msg, bank_id
            else:
                return "Fail", "Bank Add unsuccessful", search_test_status, search_msg, bank_id

        except Exception as e:
            # Handle exceptions and return failure statuses
            return (
                "Fail",
                f"Bank Add unsuccessful: {str(e)}",
                "Fail",
                "Search failed due to an exception",
                "",
            )
            
    def search(self, row_data):
        try:
            # Directly access 'Bank_Name' from the dictionary
            bank_name = row_data["Bank_Name"]
        except KeyError:
            # Handle case where the key doesn't exist
            bank_name = ""
        
        # Enter the bank name into the search field
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(bank_name)
        
        bank_id = ' '
        search_test_status, search_status, search_data = "Fail", "search data Unsuccessful", [' ']
        
        try:
            # Iterate through table rows to find the matching bank name
            for i in range(1, 6):
                row = self.driver.find_element(By.XPATH, f'//table[@id="bank_list"]//tbody//tr[{i}]')
                search_data = row.find_element(By.XPATH, './td[2]').text
                bank_id = row.find_element(By.XPATH, './td[1]').text
                if search_data == bank_name:
                    search_test_status, search_status = "Pass", "search data successful"
                    break
                else:
                    search_test_status, search_status = "Fail", "search data Unsuccessful"
        except:
            search_data=[' ']
            search_test_status = "Fail"
            search_status=("search data Unsuccessfull")
        
        return search_test_status, search_data, search_status, bank_id
    
    def edit_bank(self, row_data, bank_id):
        try:
            # Check if bank_id is empty (not found)
            if bank_id == ' ':
                Edit_test_status = "Fail"
                Edit_status = "Edit bank Unsuccessful: Bank ID not found"
                return Edit_test_status, Edit_status  # Return immediately

            # Simplified Edit check
            Edit = row_data["Edit"]

            # Check if Edit action is "yes"
            if re.match(r"yes", str(Edit), re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(bank_id)
                try:
                    # Check if the edit button exists for the row
                    edit_button = self.driver.find_element(By.XPATH, f'//table[@id="bank_list"]/tbody/tr[1]/td[4]/div/button')
                    if edit_button.is_displayed():
                        edit_button.click()
                        sleep(3)
                        self.driver.find_element(By.XPATH, '//table[@id="bank_list"]/tbody/tr[1]/td[4]/div/ul/li[1]/a').click()
                        sleep(3)
                        # Clear existing data and send new data from row_data for editing
                        self.driver.find_element(By.ID, "bank_name").clear()
                        self.driver.find_element(By.ID, "bank_name").send_keys(row_data["Edit_Data1"])
                        self.driver.find_element(By.ID, "short_code").clear()
                        self.driver.find_element(By.ID, "short_code").send_keys(row_data["Edit_Data2"])
                        self.driver.find_element(By.ID, "acc_number").clear()
                        self.driver.find_element(By.ID, "acc_number").send_keys(row_data["Edit_Data3"])
                        self.driver.find_element(By.ID, "ifsc_code").clear()
                        self.driver.find_element(By.ID, "ifsc_code").send_keys(row_data["Edit_Data4"])
                        self.driver.find_element(By.ID, "address").clear()
                        self.driver.find_element(By.ID, "address").send_keys(row_data["Edit_Data5"])
                        # Submit the form after editing
                        self.driver.find_element(By.ID, "bank_submit").click()
                        
                        Edit_test_status = "Pass"
                        Edit_status = "Edit Bank successful"
                    else:
                        Edit_test_status = "Fail"
                        Edit_status = "Edit button not found"
                except Exception as e:
                    Edit_test_status = "Fail"
                    Edit_status = f"Edit Bank Unsuccessful: {str(e)}"

            # If Edit action is "No"
            elif re.match(r"No", str(Edit), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit Bank Unsuccessful: Edit action is 'No'"

            else:
                Edit_test_status = "Fail"
                Edit_status = "Edit Bank Unsuccessful: Invalid Edit action"
            
        except Exception as e:
            Edit_test_status = "Fail"
            Edit_status = f"Edit Bank Unsuccessful: {str(e)}"
        
        return Edit_test_status, Edit_status

  
    def Delete_bank(self,row_data,bank_id):
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
                    delete_button = self.driver.find_element(By.XPATH, f'//table[@id="bank_list"]/tbody/tr[1]/td[4]/div/button')
                    if delete_button.is_displayed():
                        delete_button.click()
                        sleep(2)
                        # Click the delete option
                        self.driver.find_element(By.XPATH, '//table[@id="bank_list"]/tbody/tr[1]/td[4]/div/ul/li[2]/a').click()
                        sleep(2)
                        # Confirm deletion
                        self.driver.find_element(By.XPATH, '(//a[@onclick="delete_bank()"])').click()
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
