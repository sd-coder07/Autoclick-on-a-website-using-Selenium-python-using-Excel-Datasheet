'''from selenium import webdriver
import time

web = webdriver.Chrome()
web.get('https://payment.ivacbd.com/#application')
time.sleep(16)
sname = 'BGDDW13AB524'

name = web.find_element('xpath', '//*[@id="application"]/div[1]/div[2]/input')

name.send_keys(sname)

time.sleep(16)
'''

'''from selenium import webdriver
import time
import pandas as pd


data = pd.read_excel('C:\\Users\\dasso\\Desktop\\DataSheet.xlsx')
web = webdriver.Chrome()
web.get('https://payment.ivacbd.com/#application')
time.sleep(16)

for index, row in data.iterrows():
    # Get values from the Excel file
    web_reference = row['Webfile Number']

    web.find_element('//*[@id="application"]/div[1]/div[2]/input').send_keys(web_reference)
'''

from selenium import webdriver
from selenium.webdriver.common.by import By  # Import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from datetime import datetime, timedelta

# Load data from Excel
data = pd.read_excel(r'C:\\Users\\dasso\\Desktop\\DataSheet.xlsx')

# Create a Service object for ChromeDriver
service = Service(r'C:\\Users\\dasso\\Downloads\\chromedriver-win64\\chromedriver.exe')
web = webdriver.Chrome(service=service)

# Open the application website
web.get('https://payment.ivacbd.com/#application')
web.maximize_window()
time.sleep(10)  # Wait for the page to load

# Loop through the data from the Excel file
button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="emergencyNoticeCloseBtn"]/span'))  # Corrected XPath
    )
button.click()

for index, row in data.iterrows():
    
    web_reference = row['Webfile Number']

    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="application"]/div[1]/div[2]/input'))
    )
    input_element.send_keys(web_reference)


    mission = row['Mission']
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="highcom"]'))
    )
    input_element.send_keys(mission)

    web_reference = row['Webfile Number']

    # Use By.XPATH to locate the element correctly
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="application"]/div[1]/div[3]/input'))
    )
    input_element.send_keys(web_reference)


    IVAC_Center = row['IVAC Center']
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ivac"]'))
    )
    input_element.send_keys(IVAC_Center)

    VISA_Type = row['VISA Type']
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ivac_visa_type"]'))
    )
    input_element.send_keys(VISA_Type)


    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="application"]/div[3]/div[2]/button'))  # Corrected XPath
    )
    button.click()
    
    Full_Name = row['Full Name']
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="personal"]/div/div[1]/input'))
    )
    input_element.send_keys(Full_Name)

    email = row['eMail']
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="personal"]/div/div[2]/input'))
    )
    input_element.send_keys(email)

    contact = row['Contact#']
    input_element = WebDriverWait(web, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="person_contact_no"]'))
    )
    input_element.send_keys(contact)

    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="personal"]/div/div[4]/button'))
    )
    button.click()

    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="tos_agree"]'))
    )
    button.click()

    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="overview"]/div[2]/button'))
    )
    button.click()

    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="payment"]/div/div/div[1]/div[3]/ul/li[2]/button/img'))
    )
    button.click()


# Waitimg time
now = datetime.now()
target_time = now.replace(hour=00, minute=24, second=0, microsecond=0)

if now > target_time:
    target_time += timedelta(days=1)

time_to_wait = (target_time - now).total_seconds()
time.sleep(time_to_wait)

# Payment process
for _ in range(4):
    try:
        button = WebDriverWait(web, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="payment"]/div/div/div[2]/ul/li[4]/div/button/span'))
        )
        button.click()

        button = WebDriverWait(web, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="full-danger"]/div/div[2]/div/div/button'))
        )
        button.click()
        time.sleep(3)
    except Exception as e:
        print("Error during payment process:", e)


input("Press Enter to close the browser...")  # Keeps the window open until you press Enter

# Close the browser at the end
web.quit()


























'''#1
    for _ in range(4):
        button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="payment"]/div/div/div[2]/ul/li[4]/div/button/span'))
    )
    button.click()

    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="full-danger"]/div/div[2]/div/div/button'))
    )
    button.click()
    time.sleep(3)
    #2
    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="payment"]/div/div/div[2]/ul/li[4]/div/button/span'))
    )
    button.click()

    button = WebDriverWait(web, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="full-danger"]/div/div[2]/div/div/button'))
    )
    button.click()
    time.sleep(3)'''