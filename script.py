# Program to send bulk customized message through WhatsApp web application
# Author @inforkgodara

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time

# Load the chrome driver
from webdriver_manager.chrome import ChromeDriverManager

driver = webdriver.Chrome(ChromeDriverManager().install())
count = 0

# Open WhatsApp URL in chrome browser
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)

# Read data from excel
excel_data = pd.read_excel('Customer bulk email data.xlsx', sheet_name='Customers')

# Iterate excel rows till to finish
for column in excel_data['Name'].tolist():
    # Assign customized message
    message = excel_data['Message'][0]

    # Locate search box through x_path
    search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'

    try:
        # Wait for the search box to be clickable
        person_title = wait.until(EC.element_to_be_clickable((By.XPATH, search_box)))
    except Exception as e:
        print("Error locating search box:", e)
        break

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(str(excel_data['Contact'][count]))
    count += 1

    # Wait for 2 seconds to search contact number
    time.sleep(2)

    try:
        # Load error message in case of unavailability of contact number
        element = driver.find_element(By.XPATH, '//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        # Format the message from excel sheet
        message = message.replace('{customer_name}', column)
        person_title.send_keys(Keys.ENTER)
        actions = ActionChains(driver)
        actions.send_keys(message)
        actions.send_keys(Keys.ENTER)
        actions.perform()

# Close chrome browser
driver.quit()
