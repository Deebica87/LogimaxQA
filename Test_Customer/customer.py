from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from time import sleep
from Utils.Excel import ExcelUtils
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import re

FILE_PATH = ExcelUtils.file_path
class CustomerAutomation:
    def __init__(self,driver):
        self.driver =driver

    def customer(self):
        try:
            function_name = "customer"
            valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
            workbook = load_workbook(FILE_PATH)
            sheet = workbook[function_name]
            sleep(10)
            # Navigating through the interface
            self.driver.find_element(By.XPATH, "//a[@class='sidebar-toggle']").click()
            sleep(5)
            self.driver.find_element(By.PARTIAL_LINK_TEXT, 'Customer').click()
            sleep(8)
            self.driver.find_element(By.XPATH, '(//span[text()="Customers"])').click()
            # Process each row
            for row_num in range(2, valid_rows):
                # Define columns and dynamically fetch their values
                data = {
                    "AddCustomer":4,
                    "Customer Type": 5,
                    "Title":6,
                    "First Name":7,
                    "Last Name":8,
                    "Company Name":9,
                    "Mobile":10,
                    "Mobile Otp":11,
                    "Change Mobile No":12,
                    "Resend Otp":13,
                    "Profession":14,
                    "E-mail":15,
                    "Phone":16,
                    "Date of Birth":17,
                    "Date of Wedding":18,
                    "Religion":19,
                    "Gender":20,
                    "address1":21,
                    "address2":22,
                    "address3":23,
                    "Pincode":24,
                    "Country":25,
                    "State":26,
                    "City":27,
                    "AddressProof":28,
                    "IDName":29,
                    "Aadhar No":30,
                    "Voter ID":31,
                    "Driving ID":32,
                    "Nominee Name":33,
                    "Relationship":34,
                    "Nominee Mobile":35, 
                    "Pan":36,
                    "PanName":37, 
                    "Edit":38,
                    "Edit Mobile":39,
                    "Edit E-mail":40,
                    "Edit AddressProof":41,
                    "Edit Aadhar No":42,
                    "Edit  Voter ID":43,
                    "Edit Driving ID":44,
                }
                row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
                print(row_data)
                # Url = LoginAutomation.url(self) 
                # Call add_Payment_Mode
                status=self.customeradd_data(row_data)
                try:
                    test_status,Actual_status = status
                    if test_status == 'Fail':
                        sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                        sheet.cell(row=row_num, column=3).value = Actual_status
                        workbook.save(FILE_PATH)
                except:    
                    Cus_test_status,Cus_status,search_test_status,search_status,Customer_id = status
                    print(status)  
                    edit = self.edit_customer(row_data,Customer_id)
                    Edit_test_status,Edit_status = edit
                    print(edit)
                    test_status = "Fail" if "Fail" in [Cus_test_status,search_test_status,Edit_test_status] else "pass"
                    Actual_status = ",".join([Cus_status, search_status,Edit_status])  
                    sheet.cell(row=row_num, column=2).value = test_status  # Write Test Status
                    sheet.cell(row=row_num, column=3).value = Actual_status  # Write Actual Status
                    workbook.save(FILE_PATH) 
            print("Classification Completed")   
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)    
        except Exception as e:
                print(f"Error during login: {e}")  
                 
    def customeradd_data(self,row_data):
        sleep(5)
        self.driver.find_element(By.ID,'add_customer').click()
        satuts = self.personal(row_data)
        print(f'{satuts}7')
        return satuts
        
    def personal(self,row_data):
        sleep(5)
        Customer_Type = row_data["Customer Type"]
        print(Customer_Type)
        if Customer_Type == "Individual":
            self.driver.find_element(By.XPATH, "//input[@type='radio' and @value='1']").click()
            self.driver.find_element(By.ID, "select2-title_select-container").click()
            Title = self.driver.find_element(By.XPATH,"//input[@role='textbox']")
            Title.send_keys(row_data["Title"])
            Title.send_keys(Keys.RETURN)
            if row_data["First Name"] != None:
                self.driver.find_element(By.ID, "firstname").send_keys(row_data["First Name"])
            else:  
                self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\First_name_fail.png") 
                return 'Fail','First Name Should be Mandatory'
            self.driver.find_element(By.ID, "lastname").send_keys(row_data["Last Name"])
        elif Customer_Type == "Company":
            self.driver.find_element(By.XPATH, "//input[@type='radio' and @value='2']").click()
            self.driver.find_element(By.ID, "select2-title_select-container").click()
            Title = self.driver.find_element(By.XPATH,"//input[@role='textbox']")
            Title.send_keys(row_data["Title"])
            Title.send_keys(Keys.RETURN)
            self.driver.find_element(By.ID, "select2-title_select-container").click()
            if row_data["Company Name"] !='':
                self.driver.find_element(By.ID, "firstname").send_keys(row_data["Company Name"])
            else:
                self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Company_name_fail.png")    
                return 'Company name is mandatory'
                
        Mobile=row_data["Mobile"]
        if Mobile != None:    
            input = self.driver.find_element(By.ID, "mobile")
            input.send_keys(str(row_data["Mobile"]))
            input.send_keys(Keys.TAB)
            sleep(5)
            # otp = (int(row_data["Mobile Otp"]))
            # if otp == int(otp):
            #     self.driver.find_element(By.ID, "otp").send_keys(str(int(row_data["Mobile Otp"])))
            #     self.driver.find_element(By.ID, "verify_otp").click()
            #     Message = self.driver.find_element(By.ID, "verified").text
            #     Actual_status = Message
            #     print(Message)
            # else:
            #     No = (row_data["Mobile Otp"])
            #     print(No)  
            #     self.driver.find_element(By.ID, "otp").send_keys(row_data["Mobile Otp"])
            #     self.driver.find_element(By.ID, "verify_otp").click()
            #     Message = self.driver.find_element(By.ID, "otp_status").text
            #     print(Message)
            #     self.driver.find_element(By.ID, "close_model").click()
            #     verified = self.driver.find_element(By.ID, "verified").text
            #     print(verified)
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Mobilenumber_fail.png")    
            return "Fail",'Mobile number is mandatory' 
        sleep(5)
        self.driver.find_element(By.XPATH, '(//b[@role="presentation"])[2]').click()
        Profession = self.driver.find_element(By.XPATH,"//input[@role='textbox']")
        Profession.send_keys(row_data["Profession"])
        Profession.send_keys(Keys.ENTER)
        if row_data["E-mail"]!="":
            self.driver.find_element(By.ID, "email").send_keys(row_data["E-mail"])
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\E-mail_fail.png") 
            return 'Fail','E-mail  is mandatory'    
        self.driver.find_element(By.NAME, "phone").send_keys(str(int(row_data["Phone"])))
        sleep(2)
        #Select Religion from Dropdown
        religion = row_data["Religion"]
        select_element = self.driver.find_element(By.ID, "religion_select")
        select = Select(select_element)
        select.select_by_visible_text(religion)
        sleep(5)
        metal_type = row_data["Gender"]
        if metal_type == "Male":
            self.driver.find_element(By.XPATH, "//input[@type='radio' and @id='gender_male']").click()
        elif metal_type == "Female":
            self.driver.find_element(By.XPATH, "//input[@type='radio' and @id='gender_female']").click()
        elif metal_type == "Others":
            self.driver.find_element(By.XPATH, "//input[@type='radio' and @id='gender_others']").click()
        if (row_data["Date of Birth"])!= None:    
            D_O_B = str(row_data["Date of Birth"])
            self.driver.find_element(By.ID, "date_of_birth").send_keys(D_O_B)
        else: 
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Date_of_birth_fail.png")    
            return 'Fail','Date of Birth  is mandatory'
        sleep(5)
        
        D_O_W = str(row_data["Date of Wedding"])
        print(D_O_W)
        self.driver.find_element(By.ID, "date_of_wed").send_keys(D_O_W)
        self.driver.find_element(By.XPATH, "(//button[@name='next'])[1]").click()
        satuts = self.Address(row_data)
        print(f'{satuts}6')
        return satuts

    def Address(self,row_data):
        sleep(4)
        if (row_data["address1"])!=None:
            self.driver.find_element(By.ID, "address1").send_keys(row_data["address1"])
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\address1_fail.png")     
            return 'Fail','Address1 is mandatory' 
        self.driver.find_element(By.ID, "address2").send_keys(row_data["address2"])
        self.driver.find_element(By.ID, "address3").send_keys(row_data["address3"])
        self.driver.find_element(By.ID, "pin_code_add").send_keys(row_data["Pincode"])
        
        Country = row_data["Country"]
        State = row_data["State"]
        City = row_data["City"]
        print(Country)
        print(State)
        print(City)
        sleep(5)
        if row_data["Country"] != None and row_data["State"] != None and row_data["City"] != None:
            sleep(5)
            self.driver.find_element(By.XPATH, '//span[@id="select2-country-container"]/span[@class="select2-selection__clear"]').click()
            # self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[2]').click()
            sleep(3)
            print("1")
            country=self.driver.find_element(By.XPATH,"//input[@role='textbox']")
            country.send_keys(Country)
            country.send_keys(Keys.ARROW_DOWN)
            country.send_keys(Keys.RETURN)
            sleep(5)
            sleep(5)
            self.driver.find_element(By.XPATH, '//span[@id="select2-state-container"]/span[@class="select2-selection__clear"]').click()
            # self.driver.find_element(By.XPATH, '(//span[@class="select2-selection__arrow"])[3]').click()
            state=self.driver.find_element(By.XPATH,'(//input[@type="search"])')
            state.send_keys(State)
            state.send_keys(Keys.ENTER)
            print(State)
            sleep(5)
            try:
                if self.driver.find_element(By.XPATH, '//span[@id="select2-city-container"]/span[@class="select2-selection__clear"]'):
                    self.driver.find_element(By.XPATH, '//span[@id="select2-city-container"]/span[@class="select2-selection__clear"]').click()
            except:
                if self.driver.find_element(By.ID, 'select2-city-container'):
                    self.driver.find_element(By.ID, 'select2-city-container').click()
            city =self.driver.find_element(By.XPATH,"//input[@role='textbox']")
            city.send_keys(City)
            city.send_keys(Keys.ENTER) 
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Country_fail.png") 
            return "Fail","Country,State,City should be mandatory field"
        AddressProof = row_data["AddressProof"]
        Name = row_data["IDName"]
        Path = r"D:\CRM\Taneira\Image"  
        Image_path = f"{Path}\\{Name}.jpg"
        print(Image_path)
        sleep(10)
        match AddressProof:
            case "Aadhar ID": 
                if  row_data["Aadhar No"] != None:
                    sleep(3)
                    self.driver.find_element(By.PARTIAL_LINK_TEXT, "AADHAR").click()
                    sleep(5)
                    self.driver.find_element(By.ID, "ad_aadhar_no").send_keys(row_data["Aadhar No"])
                    self.driver.find_element(By.ID, "ad_aadhar_front_img").send_keys(Image_path)
                    self.driver.find_element(By.ID, "ad_aadhar_back_img").send_keys(Image_path)  
                else:
                    self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Aadhar_fail.png") 
                    return 'Fail','Aadhar should be mandatory field'
            case "Votter ID":
                if row_data["Voter ID"] != None:
                    sleep(2)
                    self.driver.find_element(By.PARTIAL_LINK_TEXT, "VOTER ID").click()
                    sleep(5)
                    self.driver.find_element(By.ID, "ad_vi_no").send_keys(row_data["Voter ID"])
                    self.driver.find_element(By.ID, "ad_vi_front_img").send_keys(Image_path)
                    self.driver.find_element(By.ID, "ad_vi_back_img").send_keys(Image_path)    
                else:
                    self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\VOTER_fail.png")
                    return 'Fail','VOTER ID should be mandatory field'
            case "Driving ID" :
                if row_data["Driving ID"] != None:
                    sleep(2)
                    self.driver.find_element(By.PARTIAL_LINK_TEXT, "DRIVING LICENCE").click()
                    sleep(3)
                    self.driver.find_element(By.ID, "ad_dl_no").send_keys(row_data["Driving ID"])
                    self.driver.find_element(By.ID, "ad_dl_front_img").send_keys(Image_path)
                    self.driver.find_element(By.ID, "ad_dl_back_img").send_keys(Image_path)
                else:
                    self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\DRIVING_ID_fail.png")
                    return 'Fail','Driving ID should be mandatory field'
        print(AddressProof)       
        sleep(10) 
        self.driver.find_element(By.XPATH,'(//button[@name="next"])[2]').click()
        print(AddressProof)
        status = self.Nominee(row_data)
        print(f'{status}5')
        return status
    
    def Nominee(self,row_data): 
        sleep(10)  
        if row_data["Nominee Name"] != None:
            self.driver.find_element(By.ID, "nominee_name").send_keys(row_data["Nominee Name"])
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Nominee_fail.png")
            return 'Fail',"Nominee Name Should be mandatory"   
            
        if row_data["Relationship"] != None:
            self.driver.find_element(By.ID, "nominee_relationship").send_keys(row_data["Relationship"])
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Relationship_fail.png")
            return 'Fail',"Relationship Should be mandatory"   
        Nominee_Mobile = str(row_data["Nominee Mobile"])
        self.driver.find_element(By.ID, "nominee_mobile").send_keys(Nominee_Mobile)
        self.driver.find_element(By.XPATH, "(//button[@name='next'])[3]").click()  
        status = self.KYC(row_data)
        print(f'{status}4')
        return status
    
    def KYC(self,row_data):  
        Name = row_data["PanName"]
        Path = r"D:\CRM\Taneira\Image"  
        Image_path = f"{Path}\\{Name}.jpeg"
        print(Image_path)
        sleep(5)
        self.driver.find_element(By.ID, "sameProof_as_addr").click()
        if row_data["Pan"] != None:
            self.driver.find_element(By.PARTIAL_LINK_TEXT, "PAN").click()
            self.driver.find_element(By.ID, "pan_no").send_keys(row_data["Pan"])
            self.driver.find_element(By.ID, "pan_front_img").send_keys(Image_path)
            self.driver.find_element(By.ID, "pan_back_img").send_keys(Image_path)
            self.driver.find_element(By.ID, "kyc_submit").click()
            sleep(6)
        else:
            self.driver.get_screenshot_as_file("D:\\CRM\\Taneira\\screeshot\\Pan_fail.png")
            return 'Fail',"Pan Should be mandatory"    
        status=self.search(row_data)
        print(f'{status}2')
        search_test_status,search_data,search_status,Customer_id = status
        try:
            Customer=str(row_data["Mobile"])
            if  Customer in search_data:
                Cus_test_status = "Pass"
                Cus_status = "Customer Add successful"
            else:
                Cus_test_status = "Fail"    
                Cus_status = "Customer Add Unsuccessful"     
        except Exception as e:  # Handle exceptions
                Cus_test_status = "Fail"
                Cus_status = f"Customer Add Unsuccessful: {str(e)}"          
        status = Cus_test_status,Cus_status,search_test_status,search_status,Customer_id
        print(f'{status}3')
        return status
    
    def search(self,row_data):
        Customer=str(row_data["Mobile"])
        self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Customer)
        Customer_id=' '
        try:
            for i in range(1, 6):
                cell_xpath = f'//table[@id="customer_list"]//tbody//tr[{i}]//td[4]'
                data= self.driver.find_element(By.XPATH, cell_xpath)
                id=self.driver.find_element(By.XPATH,f'//table[@id="customer_list"]//tbody//tr[{i}]//td[1]')
                Customer_id = id.text
                search_data=data.text
                if (search_data==Customer):
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
        status = search_test_status,search_data,search_status,Customer_id
        print(status)
        return status
    def edit_customer(self,row_data,Customer_id):
        try:
        # Check if Customer_id is empty
            if Customer_id == None:
                Edit_test_status = "Fail"
                Edit_status = "Edit department Unsuccessful"
                edit =  Edit_test_status,Edit_status
            if  re.match(r"yes",str(row_data["Edit"]),re.IGNORECASE):
                self.driver.find_element(By.XPATH, '//input[@type="search"]').clear()
                self.driver.find_element(By.XPATH, '//input[@type="search"]').send_keys(Customer_id) 
                try:
                    if self.driver.find_element(By.XPATH,'(//button[@class="btn btn-default dropdown-toggle"])[1]'):
                        self.driver.find_element(By.XPATH,'(//button[@class="btn btn-default dropdown-toggle"])[1]').click()
                        sleep(2)
                        self.driver.find_element(By.XPATH, '(//a[@class="btn-edit"])[1]').click()
                        Edit_Mobile=row_data["Edit Mobile"]
                        if Edit_Mobile != None:    
                            self.driver.find_element(By.ID, "mobile").clear()
                            input = self.driver.find_element(By.ID, "mobile")
                            input.send_keys(str(row_data["Edit Mobile"]))
                            input.send_keys(Keys.TAB)
                            sleep(5)
                            otp = (int(row_data["Mobile Otp"]))
                            if otp == int(otp):
                                self.driver.find_element(By.ID, "otp").send_keys(str(int(row_data["Mobile Otp"])))
                                self.driver.find_element(By.ID, "verify_otp").click()
                                Message = self.driver.find_element(By.ID, "verified").text
                                Actual_status = Message
                                print(Message)  
                        self.driver.find_element(By.ID, "email").clear()
                        self.driver.find_element(By.ID, "email").send_keys(row_data["Edit E-mail"])    
                        self.driver.find_element(By.XPATH, "(//button[@name='next'])[1]").click()
                        EditAddressProof = row_data["Edit AddressProof"]
                        Name = row_data["IDName"]
                        Path = r"D:\CRM\Taneira\Image"  
                        Image_path = f"{Path}\\{Name}.jpg"
                        print(Image_path)
                        sleep(5)
                        match EditAddressProof:
                            case "Aadhar ID": 
                                if  row_data["Edit Aadhar No"] != None:
                                    self.driver.find_element(By.PARTIAL_LINK_TEXT, "AADHAR").click()
                                    self.driver.find_element(By.ID, "ad_aadhar_no").send_keys(row_data["Edit Aadhar No"])
                                    self.driver.find_element(By.ID, "ad_aadhar_front_img").send_keys(Image_path)
                                    self.driver.find_element(By.ID, "ad_aadhar_back_img").send_keys(Image_path)  
                                else:
                                    return 'Fail','Aadhar should be mandatory field'
                            case "Votter ID":
                                sleep(2)
                                if row_data["Edit Voter ID"] != None:
                                    self.driver.find_element(By.PARTIAL_LINK_TEXT, "VOTER ID").click()
                                    self.driver.find_element(By.ID, "ad_vi_no").send_keys(row_data["Edit Voter ID"])
                                    self.driver.find_element(By.ID, "ad_vi_front_img").send_keys(Image_path)
                                    self.driver.find_element(By.ID, "ad_vi_back_img").send_keys(Image_path)    
                                else:
                                    return 'Fail','VOTER ID should be mandatory field'
                            case "Driving ID" :
                                if row_data["Driving ID"] != None:
                                    self.driver.find_element(By.PARTIAL_LINK_TEXT, "DRIVING LICENCE").click()
                                    self.driver.find_element(By.ID, "ad_dl_no").send_keys(row_data["Edit Driving ID"])
                                    self.driver.find_element(By.ID, "ad_dl_front_img").send_keys(Image_path)
                                    self.driver.find_element(By.ID, "ad_dl_back_img").send_keys(Image_path)
                                else:
                                    return 'Fail','Driving ID should be mandatory field'
                                self.driver.find_element(By.XPATH,'(//button[@name="next"])[2]').click()
                                Edit_test_status = "Pass"
                                Edit_status = "Edit Customer successful"
                except:
                    Edit_test_status = "Fail"
                    Edit_status = "Edit Customer Unsuccessful"             
                       
            elif re.match(r"No", str(["Edit"]), re.IGNORECASE):
                Edit_test_status = "Fail"
                Edit_status = "Edit Customer Unsuccessful"   
        except:         
            Edit_test_status = "Fail"
            Edit_status = "Edit Customer Unsuccessful"  
        edit =  Edit_test_status,Edit_status
        return edit  
    
    def Customername(self):
        function_name = "customer"
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]
        sleep(10)
        for row_num in range(2, valid_rows):
            # Define columns and dynamically fetch their values
            data = {
                "First Name":7,
                "Mobile":10,
            }
            row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
            name = row_data["First Name"]
            number=row_data["Mobile"]
            
            datas = name,number
            return datas

