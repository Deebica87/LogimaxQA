
# function_name = 'login'
# # Get current date and time
# dt = datetime.datetime.now()
# path= f"D:\\CRM\\Taneira\\screeshot\\{function_name}_{dt}.png"
# #strptime('%Y%m%d%H%M%S')
# print(path)

# # Fetch all elements from the dropdown
        # se_ver_list = self.driver.find_elements(By.XPATH, '//ul[@id="ui-id-1"]//li')
        # print(se_ver_list)
        # sleep(5)
        # for element in se_ver_list:
        #     print(f"Text: {element.text}")
        #     if element.text == customer_name:  # Replace this with your `customer_name` variable
        #         print(f"Clicking on: {element.text}")
        #         element.click()
        #         break
    
        # for element in se_ver_list:
        #     if element.text == customer_name:  # Replace this with your `customer_name` variable
        #         print(f"Clicking on: {element.text}")
        #         element.click()
        #         break
import datetime
from time import sleep
ct1 = datetime.datetime.now()     
sleep(5) 
ct2 = datetime.datetime.now()    
time_diff = ct2 - ct1  
print("ct1 =", ct1.strftime("%Y-%m-%d %H:%M:%S.%f"))  # Format output
print("ct2 =", ct2.strftime("%Y-%m-%d %H:%M:%S.%f"))  
print("Time difference =", time_diff) 