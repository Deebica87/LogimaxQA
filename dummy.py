# # import pandas as pd
# # import re
# # import win32com.client
# # from openpyxl import load_workbook
# # import openpyxl
# # from Test_login.login import LoginAutomation

# # file_path = "D:\\CRM\\Taneira\\log.xlsx"  

# # df = pd.read_excel(file_path)
# # function_name = "Login"
# # workbook = load_workbook(file_path)
# #             # workbook = load_workbook(file_path)
# # sheet = workbook[function_name]
# # i =2
# # count = 0
# # while (i<100):
# #     cellvalue = sheet.cell(row =i, column=1).value
# #     if cellvalue is None:
# #         break
# #     i=i+1
# #     count = count+1
        

# # print(count)

# # # Define the variables
# # a = "pass"
# # b = "pass"
# # c = "pass"
# # d = "pass"
# # e = "fail"

# # # Check the status of all variables
# # if "fail" in [a, b, c, d, e]:
# #   output = "fail"  
# # else :
# #   "pass"

# # print(output)
# # a = "pass"
# # b = "pass"
# # c = "pass"
# # d = "pass"

# # # Combine the variable values into a comma-separated string

# # cellvalues = ["Pass", "PAss", "pASs", "Fail", "pass", "FAIL", "Pass"]

# # # Count occurrences directly
# # Pass = cellvalues.count("Pass")  # Case-sensitive count for "Pass"
# # Fail = cellvalues.count("Fail")  # Case-sensitive count for "Fail"

# # print(f"Pass: {Pass}, Fail: {Fail}")


# # a = 1
# # try:
# #   if a == '':
# #       Edit_test_status = "Fail"
# #       Edit_status = "Edit department Unsuccessful"   
# #   elif a==1:
# #     Edit_test_status = "pass"
# #     Edit_status = "Edit department successful"
# # except:
# #   pass
# # edit = Edit_test_status,Edit_status 
# # print(edit) 

# # url = 'https://php8.logimaxindia.com/titan_v7/admin/index.php/admin/dashboard'

# # # Split the URL from the last '/'
# # base_url = url.split('php/')[0]+'php/pp/common/payment_mode/list'
# # # (url[0])
# # print(base_url)

# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from time import sleep
# from Utils.Excel import ExcelUtils
# from openpyxl import load_workbook
# import re

# FILE_PATH = "D:\CRM\Taneira\log.xlsx"
# class  BankAutomation:
#     def __init__(self, driver):
#         self.driver = driver
        
#     def  bank(self):
#         try:    
#             function_name = "Bank"
#             valid_rows = ExcelUtils.get_valid_rows(FILE_PATH,function_name)
#             workbook = load_workbook(FILE_PATH)
#             sheet = workbook[function_name]
#             sleep(10)
#             self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
#             sleep(5)
#             self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Masters').click()
#             self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#             sleep(8)
#             self.driver.find_element(By.XPATH, '(//span[text()="Bank"])').click()
#             for row_num in range(2,valid_rows):
#                 Bank_Name = sheet.cell(row=row_num, column=4).value
#                 print(Bank_Name)
#                 short_code = sheet.cell(row=row_num, column=5).value
#                 print(short_code)
#                 Acc_NO = sheet.cell(row=row_num, column=6).value
#                 print(Acc_NO)
#                 IFSC = sheet.cell(row=row_num, column=7).value
#                 print(IFSC)
#                 Bank_Address = sheet.cell(row=row_num, column=8).value
#                 print(Bank_Address)
#                 Edit = sheet.cell(row=row_num, column=9).value
#                 print(Edit)
#                 Edit_data1 = sheet.cell(row=row_num, column=10).value
#                 print(Edit_data1)
#                 Edit_data2 = sheet.cell(row=row_num, column=11).value
#                 print(Edit_data2)
#                 Edit_data3 = sheet.cell(row=row_num, column=12).value
#                 print(Edit_data3)
#                 Edit_data4 = sheet.cell(row=row_num, column=13).value
#                 print(Edit_data4)
#                 Edit_data5 = sheet.cell(row=row_num, column=14).value
#                 print(Edit_data5)
#                 Delete = sheet.cell(row=row_num, column=15).value
#                 print(Delete)
#                 self.driver.refresh()
#                 # Call add_Payment_Mode
#                 status=self.add_Payment_Mode(Bank_Name,short_code,Acc_NO,IFSC,Bank_Address)
#                 Desi_test_status,Desi_status,search_test_status,search_status,pay_id = status
#                 print(status)
#                 edit = self.edit_Payment_Mode(Edit,Edit_data1,Edit_data2,pay_id)
#                 Edit_test_status,Edit_status = edit   
#                 print(edit)
#                 sleep(3)
#                 delete = self.delete(Delete,pay_id)
#                 Delete_test_status,Delete_status = delete
#                 print(delete)
#                 test_status = "Fail" if "Fail" in [Desi_test_status,search_test_status,Edit_test_status, Delete_test_status] else "pass"
#                 Actual_status = ",".join([Desi_status, search_status, Edit_status, Delete_status])  
#                 sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
#                 sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
#                 workbook.save(FILE_PATH)
#             print("Payment Mode Completed")   
#             Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
#             print(Status)
#             Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)         
#         except Exception as e:
#                 print(f"Error during login: {e}") 
                        
#     def add_Payment_Mode(self,Bank_Name,short_code,Acc_NO,IFSC,Bank_Address):
#         sleep(2)
#         self.driver.find_element(By.ID, 'add_bnk').click()   
#         sleep(3)
#         self.driver.find_element(By.ID, "bank_name").send_keys(Bank_Name)
#         self.driver.find_element(By.ID, "short_code").send_keys(short_code)
#         self.driver.find_element(By.ID, "acc_number").send_keys(Acc_NO)
#         self.driver.find_element(By.ID, "ifsc_code").send_keys(IFSC)
#         self.driver.find_element(By.ID, "address").send_keys(Bank_Address)
#         sleep(5)
#         self.driver.find_element(By.XPATH, '(//button[@class="btn btn-primary"])[1]').click()
#         # url = LoginAutomation.perform_login(url)
#         # base_url = url.split('php/')[0]+'php/pp/common/payment_mode/list'
#         self.driver.get("https://php8.logimaxindia.com/titan_v7/admin/index.php/pp/common/payment_mode/list")
#         try:
#             # if self.driver.title == "Payment Mode - Payment_mode | TANEIRA":
#                 sleep(5)
#                 status=self.search(Bank_Name)
#                 search_test_status,search_data,search_status,pay_id = status
#                 if Bank_Name in search_data:
#                     Desi_test_status = "Pass"
#                     Desi_status = "Payment Mode Add successful"
#                 else:
#                     Desi_test_status = "Fail"    
#                     Desi_status = "Payment Mode Add Unsuccessful"   
#             # else:       
#             #     Desi_test_status = "Fail"    
#             #     Desi_status = "Payment Mode Add Unsuccessful"
#             #     search_test_status = "Fail" 
#             #     search_status = "Data not added"
#             #     pay_id = ' '
#         # Handle exceptions              
#         except Exception as e: 
#             Desi_test_status = "Fail"
#             Desi_status = f"Payment Mode Add Unsuccessful:{str(e)}"          
#         status = Desi_test_status,Desi_status,search_test_status,search_status,pay_id
#         return status
            
#     def search(self,Payment_Mode):
#         self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Payment_Mode)
#         pay_id=' '
#         try:
#             for i in range(1, 6):
#                 cell_xpath = f'//table[@id="paymode_list"]//tbody//tr[{i}]//td[2]'
#                 data= self.driver.find_element(By.XPATH, cell_xpath)
#                 id=self.driver.find_element(By.XPATH,f'//table[@id="paymode_list"]//tbody//tr[{i}]//td[1]')
#                 pay_id = id.text
#                 search_data=data.text
#                 if (search_data==Payment_Mode):
#                     search_test_status = "Pass"    
#                     search_status=("search data successfull")
#                     break
#                 else:
#                     search_test_status = "Fail"
#                     search_status=("search data Unsuccessfull") 
#         except:
#             search_data=[' ']
#             search_test_status = "Fail"
#             search_status=("search data Unsuccessfull")
#         status = search_test_status,search_data,search_status,pay_id
#         return status    
#     def edit_Payment_Mode(self,Edit,Edit_data1,Edit_data2,pay_id):
#         try:
#         # Check if Dept_id is empty
#             if pay_id == ' ':
#                 Edit_test_status = "Fail"
#                 Edit_status = "Edit Payment Mode Unsuccessful"
#                 edit =  Edit_test_status,Edit_status
#                 return edit  # Return immediately
#         # Check if Edit action is "yes" (case-insensitive)
#             if re.match(r"yes", str(Edit), re.IGNORECASE):
#                 self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
#                 self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(pay_id)
#                 try:   
#                     if  self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button'):
#                         self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button').click()
#                         sleep(3)
#                         self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/ul/li[1]/a').click()
#                         sleep(3)
#                         self.driver.find_element(By.ID,"mode_name").clear()
#                         self.driver.find_element(By.ID,"mode_name").send_keys(Edit_data1)
#                         self.driver.find_element(By.ID,"short_code").clear()
#                         self.driver.find_element(By.ID,"short_code").send_keys(Edit_data2)
#                         self.driver.find_element(By.ID,"paymode_submit").click()
#                         Edit_test_status = "Pass"
#                         Edit_status = "Edit Payment Mode successful"
#                 except:
#                     Edit_test_status = "Fail"
#                     Edit_status = "Edit Payment Mode Unsuccessful"       
#             elif re.match(r"No", str(Edit), re.IGNORECASE):
#                 Edit_test_status = "Fail"
#                 Edit_status = "Edit Payment Mode Unsuccessful"   
#         except:         
#             Edit_test_status = "Fail"
#             Edit_status = "Edit Payment Mode Unsuccessful"  
#         edit =  Edit_test_status,Edit_status
#         return edit  
            
#     def delete(self,Delete,pay_id):
#         try:
#         # Check if Dept_id is empty
#             if pay_id == ' ':
#                 Edit_test_status = "Fail"
#                 Edit_status = "Delete Payment Mode Unsuccessful"
#                 edit =  Edit_test_status,Edit_status
#                 return edit  # Return immediately
#         # Check if Edit action is "yes" (case-insensitive)
#             if re.match(r"yes", str(Delete), re.IGNORECASE): 
#                 self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
#                 self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(pay_id) 
#                 try:
#                     if self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button'):
#                         self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/button').click()
#                         sleep(2)
#                         self.driver.find_element(By.XPATH,'//table[@id="paymode_list"]/tbody/tr[1]/td[5]/div/ul/li[2]/a').click()
#                         sleep(2)
#                         self.driver.find_element(By.XPATH,'(//a[@onclick="deletePaymode(id,2)"])').click()
#                         sleep(2)
#                         Delete_test_status = "Pass"
#                         Delete_status = "Delete Payment Mode  successful"
#                 except:
#                     Delete_test_status = "Fail"
#                     Delete_status = "Delete Payment Mode  Unsuccessful"         
#             elif re.match(r"No", str(Delete), re.IGNORECASE):
#                 Delete_test_status = "Fail"
#                 Delete_status = "Delete Payment Mode  Unsuccessful" 
#             delete = Delete_test_status,Delete_status
#             return delete
#         except: 
#             Delete_test_status = "Fail"
#             Delete_status = "Delete Payment Mode  Unsuccessful" 
#         delete = Delete_test_status,Delete_status
#         return delete
    
    
# # import random
# import random

# # prints a random value from the list
# list1 = [1, 2, 3, 4, 5, 6]
# print(random.choice(list1))    

# a = r"D:\CRM\Taneira\Image"  # Use a raw string to handle backslashes
# b = "Aadhar ID"

# # Combine the strings and add the .jpg extension
# output = f"{a}\\{b}.jpg"

# print(output)
a=[1,1,1]

for i in a:
    print(i)
