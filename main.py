import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


df = pd.read_excel("challenge.xlsx")
df.rename(columns={'Phone Number': 'Phone', 'Company Name': 'CompanyName', 'Role in Company': 'Role',
                   'Last Name ': 'LastName', 'First Name': 'FirstName'}, inplace=True)
excel_rows = df.to_dict('records')

s = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()

options.add_argument('headless')
options.add_argument('disable-extensions')
options.page_load_strategy = 'eager'
options.add_argument("--incognito")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=s, options=options)
rpa_website = driver.get("http://www.rpachallenge.com/")


def css_data_input(row):
    first_name = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelFirstName"]')
    last_name = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelLastName"]')
    company_name = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelCompanyName"]')
    role = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelRole"]')
    address = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelAddress"]')
    email = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelEmail"]')
    phone = driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelPhone"]')
    submit_button = driver.find_element(By.CLASS_NAME, 'btn.uiColorButton')

    first_name.click()
    first_name.send_keys(row['FirstName'])
    last_name.click()
    last_name.send_keys(row['LastName'])
    company_name.click()
    company_name.send_keys(row['CompanyName'])
    role.click()
    role.send_keys(row['Role'])
    address.click()
    address.send_keys(row['Address'])
    email.click()
    email.send_keys(row['Email'])
    phone.click()
    phone.send_keys(row['Phone'])
    submit_button.click()


start_button = driver.find_element(By.CLASS_NAME, 'btn-large.uiColorButton')
start_button.click()

for row in excel_rows:
    css_data_input(row)

message = driver.find_element(By.CSS_SELECTOR, value='div.message2')
a = message.get_attribute('innerText')

reset = driver.find_element(By.CSS_SELECTOR, value='.btn-large.uiColorButton')
reset.click()

print("Задача выполнена. Итоговый результат: " + "\n" + a)
