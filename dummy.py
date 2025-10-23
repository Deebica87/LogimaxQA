# import re

# def validate_input(val):
    #Allow only letters/spaces, minimum 3 characters
    # pattern = r"^[A-Za-z\s]{3,}$"
    # #Allow exactly 10 digits (mobile number)
    # pattern = r"^\d{10}$"   # ✅ now it's a string, not a tuple
    # pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    # pattern = r"^[\w\s,./-]{5,}$"
    # pattern=r"^\d{6}$" 
#     pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1}$"
    
#     if val == "" or re.fullmatch(pattern, val):
#         print(f"'{val}' → Accepted ✅")
#         return True
#     else:
#         print(f"'{val}' → Not accepted ❌")
#         return False
    
# for test in ['Text','S','988','@^*^@','89.909','67.85','tes.rom', 'ty432','$%^^^TFV','Sri@gmail.com',
#              'sri@gmail.co.in','v@India.co.in','12/6,sangam street coimbatore.','643890', "29ABCDE1234F1Z5",  
#             "07ABCDE1234F1Z2","ABCDE1234F1Z5","29ABCDE1234F1A5", "29abcde1234f1z5"]:
#     validate_input(test)
    
# PATTERNS = {
#     "GrossWt": r"^\d+(\.\d{1,3})?$",     # numbers, up to 3 decimal places
#     "firstName": r"^[A-Za-z\s]+$",       # only letters & spaces
#     "lastName": r"^[A-Za-z\s]+$",        # same rule as first name
#     "mobileNumber": r"^\d{10}$",         # exactly 10 digits
#     "email": r"^[\w\.-]+@[\w\.-]+\.\w{2,}$",  # simple email validation
#     "address1": r"^[A-Za-z0-9\s,.-]+$",  # letters, digits, spaces, , . -
#     "address2": r"^[A-Za-z0-9\s,.-]*$",  # optional address
#     "address3": r"^[A-Za-z0-9\s,.-]*$",  # optional address
#     "pincode": r"^\d{6}$"                # Indian 6-digit pincode
# }





# a = ["GrossWt","Pcs","Rate"]

# if any(item in a for item in ["GrossWt","Pcs","Rate"]):
#     print("Yes")
# else:
#     print("No")


# # a = ["GrossWt","Pcs","Rate"]


# for data in ["GrossWt", "Pcs", "Rate","Mcvalue","wastage percentage"]:
#     if data in a:
#         print(data)
#     else:
#         print("No")

# a = ["GrossWt", "Pcs", "Rate","Mcvalue"]
# remove_items = ["wastage percentage", "Mcvalue"]

# y=[]
# for x in a:
#     if x not in remove_items:
#         print(x)
#         y.append(x)
#     else:
#         print('no')
# print(y)
# a = [x for x in a if x not in remove_items]
# print(a)

# import openpyxl
# from openpyxl.drawing.image import Image
# from selenium import webdriver
# import time

# # --- Step 1: Take Screenshot ---
# driver = webdriver.Chrome()
# driver.get("https://www.google.com")   # example page
# time.sleep(2)

# screenshot_path = "page_screenshot.png"
# driver.save_screenshot(screenshot_path)
# driver.quit()

# # --- Step 2: Create Excel & Insert Screenshot ---
# excel_path = "Test_Report.xlsx"

# # Create workbook
# workbook = openpyxl.Workbook()
# sheet = workbook.active
# sheet.title = "Report"

# # Write some text
# sheet["A1"] = "Test Case Screenshot"

# # Insert the image into cell B2
# img = Image(screenshot_path)
# sheet.add_image(img, "B2")

# # Save Excel file
# workbook.save(excel_path)

# print(f"✅ Screenshot saved in Excel: {excel_path}")
# value=8939324032	
# Next_No =str(value)

# Cus_No='8939324032'

# if str(Next_No) == Cus_No:
#     print(Next_No)
#     print(type(Next_No))
# else:
#     False


# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import Select
# from selenium.common.exceptions import NoSuchElementException
# from selenium.common.exceptions import NoAlertPresentException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from time import  sleep
# import unittest
# from Utils.Excel import ExcelUtils
# from openpyxl.drawing.image import Image
# from openpyxl import load_workbook
# from openpyxl.styles import Font
# from Utils.Function import Function_Call
# import re


# FILE_PATH = ExcelUtils.file_path
# class Contract_Price(unittest.TestCase):
#     def __init__(self,driver):
#         self.driver =driver   
#         self.wait = WebDriverWait(driver, 30)

#     def test_contractprice(self,test_case_id):
#         driver = self.driver
#         wait = self.wait
#         sleep(3)
#         function_XPATH = 'Contract_Price'
#         test_case_id = test_case_id
#         value =ExcelUtils.Test_case_id_count(FILE_PATH, function_XPATH,test_case_id)
#         print(value)
#         valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_XPATH)
#         workbook = load_workbook(FILE_PATH)
#         sheet = workbook[function_XPATH]
#         row=1
#         count = value
#         for row_num in range(2, valid_rows):
#             current_id = sheet.cell(row=row_num, column=1).value  # Column 1 = Test Case Id
#             if current_id == test_case_id:
#                 data = {
#                     "Mc": 16,
#                     "Charges": 17,
#                     "Stone Calc Type": 18,
#                     "UOM": 19,
#                     "Quality Code": 20,
#                     "From Cent": 21,
#                     "To Cent": 22,
#                     "Rate": 23,
#                     "Image": 24
#                 }

#                 row_data = {
#                     key: sheet.cell(row=row_num, column=col).value
#                     for key, col in data.items()
#                 }
#                 print(row_data)
#                 Mandatory_field=[]
#                 # MC_Type 
#                 if row_data["MC_Type"] is not None:
#                     Function_Call.dropdown_select(f'//*[@id="select2-o_item[{row}][va_type]-uz-container"]', row_data["MC_Type"],"/html/body/span/span/span[1]/input")
#                 else:
#                     msg = f"'{None}' → MC_Type field is mandatory ⚠️"
#                     Mandatory_field.append(msg); print(msg); self.Remark(row_num, msg)
                    
#                 Error_field_val=[]    
                
#                 if row_data["Mc"]is not None:
#                     errors=Function_Call.fill_input(
#                         wait,
#                         locator=(By.XPATH, f'//input[@name="o_item[{row}][mc]"]'),
#                         value=row_data["Mc"],
#                         #Allow exactly 10 digits (mobile number)
#                         pattern=r"^\d+(\.\d{1,3})?$",
#                         field_name="Mc",
#                         screenshot_prefix="Mc",
#                         range_check = lambda v: 0 <= float(v) <= 100,
#                         row_num=row_num
#                     )
#                     Error_field_val.extend(errors)
#                     print(Error_field_val)
#                 else:
#                     msg = f"'{None}' → Mc field is mandatory ⚠️"
#                     Mandatory_field.append(msg); print(msg); self.Remark(row_num, msg)
                    
                    
                    
                    
#         def Remark(self,row_num,Field_validation_satus): 
#             # Load the workbook
#             workbook = load_workbook(FILE_PATH)
#             sheet = workbook.active  # or workbook["SheetName"]
#             if Field_validation_satus:
#                 sheet.cell(row=row_num, column=31, value=Field_validation_satus).font = Font(bold=True, color="FF8000")
#             # Save workbook
# #             workbook.save(FILE_PATH)


# from openpyxl import load_workbook

# # Load the Excel file
# file_path = "D:\Retail_Testing\Retail_data.xlsx"
# Sheet_name = 'Contract_Price'
# workbook = load_workbook(file_path)
# sheet = workbook[Sheet_name] # or workbook["Sheet1"]

# # Find the column number of "Status"
# status_col_num = None
# for col in range(1, sheet.max_column + 1):
#     cell_value = sheet.cell(row=1, column=col).value
#     if cell_value and cell_value.strip().lower() == "field_validation_status":
#         status_col_num = col
#         print(type(status_col_num))  # <class 'int'>
#         break
   
# print(f"Status column number is: {status_col_num}")

# count =1
# if count > 1:
#     count -= 1
#     print('yes')
# else:
#     print('No')
    
    
# import re

# pattern = r"^\d{9,18}$"

# tests = ["123456789", "987654321012", "123456789012345678", "12345", "12345abcd","^*^*&*^*^*","12345678901234567890"]

# for t in tests:
#     if re.fullmatch(pattern, t):
#         print(f"'{t}' ✅ Accepted")
#     else:
#         print(f"'{t}' ❌ Not Accepted")
        
# import re
# pan_pattern = r"^[A-Z]{5}[0-9]{4}[A-Z]$"
# aadhaar_pattern = r"^[2-9]{1}[0-9]{11}$"
# gst_pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1}$"
# cin_pattern = r"^[LU]{1}[0-9A-Z]{5}[0-9]{2}[0-9]{4}[A-Z]{3}[0-9]{6}$"


# def validate(val, pattern, name):
#     if re.fullmatch(pattern, val):
#         print(f"{name}: '{val}' ✅ Valid")
#     else:
#         print(f"{name}: '{val}' ❌ Invalid")

# validate("ABCDE1234F", pan_pattern, "PAN")
# validate("234567890123", aadhaar_pattern, "Aadhaar")
# validate("29ABCDE1234F1Z5", gst_pattern, "GST")
# validate("L12345MH2020PLC123456", cin_pattern, "CIN")

# count=1
# if count > 1:
#     print('yes')
# else:
#     print("no")
    
#     import re

# pattern = r"^[2-9]\d{11}$"

# samples = ["123456789012", "012345678901", "234567890123", "987654321000","213412341234"]

# for s in samples:
#     if re.fullmatch(pattern, s):
#         print(f"{s} ✅ Valid Aadhaar")
#     else:
#         print(f"{s} ❌ Invalid Aadhaar")
        
# import re

# pattern = r"^[A-Z]{4}0\d{6}$"

# samples = ["SBIN0001234", "HDFC0005678", "abcd0123456", "ICIC0000001", "SBIN0001234"]

# for s in samples:
#     if re.fullmatch(pattern, s):
#         print(f"{s} ✅ Valid IFSC")
#     else:
#         print(f"{s} ❌ Invalid IFSC")
        
# ImagesFront = "Aadhar JPG"
# Path = r"D:\Retail_Testing\Image_all_Format"
# Image_path = f"{Path}\{ImagesFront}"
# print (Image_path)


# a ="Aadhar jpg" 
# b ="Aadhar GIF" 
# c ="Aadhar BMP" 
# d ="Aadhar png"


# if "JPG" in a.upper() or "PNG" in a.upper():
#     print(a)

# if "JPG" in b or "PNG" in b:
#     print(b)

# if "JPG" in c or "PNG" in c:
#     print(c)

# if "JPG" in d or "PNG" in d:
#     print(d)

# row=2

# if row > 1:
#  print("yefs")
# else:
#   print('nod')
  



# gross_weight=12
# dust_weight =0.100
# Stone_weight = 1.200
# net_weight = 10.602
# wastage_percentage = 0.92
# Gold_Rate = 7105
# Exchangevalue =98
# Stone_rate = 1500

# import math

# Find_Wt= gross_weight-dust_weight-Stone_weight
# Find_Wt=round(Find_Wt, 3)
# Va = (wastage_percentage/100)*Find_Wt
# Va = round(Va, 3)
# Find_NWt = Find_Wt - Va
# Find_NWt=round(Find_NWt, 3)

# if Find_NWt == net_weight:
#     val=(net_weight*Gold_Rate)/100*Exchangevalue
#     Amt = val+Stone_rate
#     Amt=("{:.2f}".format(math.ceil(Amt)))
#     print(Amt)
# else:
#     print('Net Value Not match')
    


# import math

# gross_weight = 12
# dust_weight = 1
# Stone_weight = 0.040
# net_weight = 10.850
# wastage_percentage = 1
# Gold_Rate = 7500
# Exchangevalue = 100
# Stone_rate = 88000

# # Calculate weights
# Find_Wt = gross_weight - dust_weight - Stone_weight
# Find_Wt = round(Find_Wt, 3)

# Va = (wastage_percentage / 100) * Find_Wt
# Va = round(Va, 3)

# Find_NWt = Find_Wt - Va
# Find_NWt = round(Find_NWt, 3)

# # Calculate amount
# if Find_NWt == net_weight:
#     val = (net_weight * Gold_Rate) / 100 * Exchangevalue
#     Amt = val + Stone_rate
#     # Round up to 2 decimal places and keep as float
#     value = math.ceil(Amt * 100) / 100
#     print(f"cal amount {value:.2f}")  # prints 169375.00
# else:
#     print("Net Value Not match")
#     value = 0.0  # or whatever default you want

# # value is float with 2 decimal precision
# print (value)
# print(type(value))  # <class 'float'>




Call_Tag = ('43700.84', 'Pass', '✅ Calculation Value is correct 43700.84')
Call_Non_Tag = ('23735.32', 'Pass', '✅ Calculation Value is correct 23735.32')
Call_HomeBill = ('188183.06', 'Fail', '❌ Calculation Error in 188183.06 | Web Value=188182.13')

Actual_Status = []
Total_amount = []

if Call_Tag:
    print(Call_Tag)
    Amount, Test_Status, Status = Call_Tag
    Total_amount.append(Amount)
    Actual_Status.append(Status)

if Call_Non_Tag:
    print(Call_Non_Tag)
    Amount, Test_Status, Status = Call_Non_Tag
    Total_amount.append(Amount)
    Actual_Status.append(Status)

if Call_HomeBill:
    print(Call_HomeBill)
    Amount, Test_Status, Status = Call_HomeBill
    Total_amount.append(Amount)
    Actual_Status.append(Status)

print("Total_amount list:", Total_amount)

total = sum(float(v) for v in Total_amount)
print("Total =", total)

print("Actual_Status list:", Actual_Status)

import re

row_values = ['9783', 'GBT-01414', 'GOLD BANGLES-18069', 'KERALA', 'LK FANCY BANGLE', 
              'Mc & Wast On Gross', '1', '2.000', '0.000', '2.000', '30.00', '30.00']

# Clean product name (3rd element → index 2)
row_values[2] = re.sub(r'-\d+$', '', row_values[2]).strip()

print(row_values)

