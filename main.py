from selenium import webdriver
from Utils.Excel import ExcelUtils
from Test_login.login import LoginAutomation
from Test_Master import Department,Designation,Profession,Payment_mode,Bank,Classification
from Test_Customer import customer
from Test_Employee import employee
from Test_Accounts import Create_Account
import datetime

class main():
    def main():
        FILE_PATH = ExcelUtils.file_path
        # Step 1: Initialize WebDriver
        ct1 = datetime.datetime.now()
        driver = webdriver.Chrome()  # Ensure ChromeDriver is in PATH
        driver.maximize_window()

        try:
            sheet_names = ExcelUtils.get_sheet_names(FILE_PATH)
            print(sheet_names)
            # Step 2: Get functions to execute from the Master sheet
            functions_to_execute = ExcelUtils.get_master_sheet_data(FILE_PATH)
            print(functions_to_execute)
        
            for function_name in functions_to_execute:
                print("funct",function_name)
                if function_name in sheet_names:
                    print("sheet",sheet_names)
                    # Initialize LoginAutomation
                    match function_name:
                        case "Login":
                            print("yes")
                            login_automation = LoginAutomation(driver)
                            Data = login_automation.perform_login()
                        case "Department":
                            Department_automation = Department.DepartmentAutomation(driver)
                            data = Department_automation.Department()
                        case "Designation":    
                            Designation_automation = Designation.DesignationAutomation(driver)
                            data = Designation_automation.Designation()
                        case "Profession":  
                            Profession_automation = Profession.ProfessionAutomation(driver)
                            data = Profession_automation.Profession()  
                        case "Payment Mode":    
                            PaymentMode_automation = Payment_mode.PaymentModeAutomation(driver)
                            data = PaymentMode_automation.Payment_Mode() 
                        case "Bank": 
                            Bank_automation = Bank.BankAutomation(driver)
                            data = Bank_automation.bank() 
                        case "classification": 
                            classification_automation = Classification.ClassificationAutomation(driver)
                            data =classification_automation.classification()   
                        case "customer": 
                            customer_automation = customer.CustomerAutomation(driver)
                            data =customer_automation.customer() 
                        case "Employees": 
                            Employees_automation = employee.EmployeesAutomation(driver)
                            data =Employees_automation.Employees() 
                        case "CreateAccount":
                            CreateAccount_automation = Create_Account.CreateAccountAutomation(driver)
                            data =CreateAccount_automation.CreateAccount()                 
                        case _:
                            print("Invalid options") # Default case                  
                else:
                    print("Invalid option")
        finally:
            # Close the WebDriver
            
            driver.close()
            driver.quit()
            print("Automation process completed.")
            ct2 = datetime.datetime.now()
            time_diff = ct2 - ct1  
            print("ct1 =", ct1.strftime("%Y-%m-%d %H:%M:%S.%f"))  # Format output
            print("ct2 =", ct2.strftime("%Y-%m-%d %H:%M:%S.%f"))  
            print("Time difference =", time_diff) 

    if __name__ == "__main__":
        print(__name__)
        main()
        
        
