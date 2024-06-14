from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
import time
import pandas as pd
from flask import Flask
import os

app = Flask(__name__)

@app.route('/')
def index():
    send_whatsapp_messages()
    return 'Messages sent successfully'

def send_whatsapp_messages():
    excel_file = 'contacts.xlsx'
    if not os.path.exists(excel_file):
        print("Excel file 'contacts.xlsx' not found in the root directory.")
        return

    df = pd.read_excel(excel_file)
    
   
    chrome_driver_path = '/path/to/your/chromedriver'  
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized") 
    driver = webdriver.Chrome(options=options)
    
    driver.get('https://web.whatsapp.com')  

    print("Please scan the QR code to log in to WhatsApp Web.")
    input("Press Enter after scanning the QR code")

    generic_message = "Hello, this is a test message from our bulk sender."

    for index, row in df.iterrows():
        name = row['Name']
        number = row['Number']
        message = f"Hi {name}, {generic_message}"
        url = f"https://web.whatsapp.com/send?phone={number}&text={message}"
        
        driver.get(url)
        time.sleep(15) 

        try:
            send_button = driver.find_element("xpath", "//button[@data-testid='send']")
            send_button.click()
            time.sleep(10) 
            print(f"Message sent to {name} at number {number}.")
        except NoSuchElementException as e:
            print(f"Failed to find 'Send' button: {e}")
        except ElementNotInteractableException as e:
            print(f"Failed to click 'Send' button: {e}")
        except Exception as e:
            print(f"Failed to send message to {name} at number {number}: {e}")
    
    driver.quit()

if __name__ == '__main__':
    app.run(port=5000, debug=True)
