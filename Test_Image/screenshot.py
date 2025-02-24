import datetime

class image:            
    #Get screenshot
    def screenshot(function_name):
        dt =datetime.datetime.now().strftime("%Y.%m.%d_%H.%M")
        path= f"D:\\CRM\\Taneira\\screeshot\\{function_name}_{dt}.png"
        return(path)