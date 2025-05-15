import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from time import sleep
import pandas as pd
import os
import datetime
import time
from openpyxl import load_workbook , Workbook
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Open chrome with current profile
options = Options()

# Path to your chrome profile
chrome_profile_path = "C:\\Users\\Minh Nhat\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1"
options.add_argument("user-data-dir=" + chrome_profile_path)

# Disable automation extension
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches",["enable-automation"])
options.add_argument("--disable-blink-features=AutomationControlled") 

# Path to your chrome driver
service = Service(r"C:\Program Files\Google\chromedriver-win32\chromedriver-win32\chromedriver.exe")
driver = webdriver.Chrome(service=service,options=options)

# Now, any Chrome window opened by Selenium will use the specified profile
driver.get("https://login.taobao.com/member/login.jhtml")

# Login to TaoBao
username_input = driver.find_element(By.NAME, "fm-login-id")
username_input.send_keys('id')
password_input = driver.find_element(By.NAME, "fm-login-password")
password_input.send_keys('password')
login_button = driver.find_element(By.XPATH,"//button[text()='登录']") 
login_button.click()

time.sleep(8)

# find element of homepage
homepage_button = driver.find_element(By.XPATH,"//span[text()='淘宝网首页']")
homepage_button.click()

time.sleep(8)

# Find product names and links
title = []
prices = []  
img_urls = []
Sold_num = []
Shop_name =[]
Shop_link = []
Link_of_product = []

# List of Keywords to search:
list_keys = ["冬衣","女式 T 恤", "裙子"] # DS tên sản phẩm muốn lấy, dùng gg dịch

# Num of pages every product to collect data
num_clicks = 2  # input num_pages of product u want to get 

# Create loop of "Next" button
def scroll_with_speed(scroll_speed):
    current_scroll_position = 0
    page_height = driver.execute_script("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
    # Scroll down the bottom of page to load data
    while current_scroll_position < page_height:
        driver.execute_script(f"window.scrollBy(0, {scroll_speed});")
        current_scroll_position += scroll_speed
        time.sleep(0.3) 
    # Scroll up to find "next" button
    driver.execute_script("window.scrollBy(0, -900);")
    
for _ in list_keys:
    searchbar = driver.find_element(By.CSS_SELECTOR,"input.rax-textinput.rax-textinput-placeholder-6.searchbar-input")
    searchbar.send_keys(_) # input Keyword 
    search_button = driver.find_element(By.XPATH,"//span[text()='搜索']")
    search_button.click() # click to find product 
    
    # Best seller button
    bestseller_button = driver.find_element(By.XPATH,"//div[text()='销量']")
    bestseller_button.click()
    time.sleep(3)
    
    # Start running loop
    for _ in range(num_clicks):
        scroll_with_speed(50)
        time.sleep(5)
        
        # Extract all elements by soup
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Extract Title of product
        elements = driver.find_elements(By.XPATH, "//span[@class='']")
        title.extend([element.text for element in elements])

        # Get image
        products = soup.find_all('div', class_='MainPic--mainPicWrapper--iv9Yv90')
        img_urls.extend([product.find('img')['src'] if product.find('img') else None for product in products])
        img_urls = [value for value in img_urls if value is not None]

        # Find all elements of Price with the class 
        price_elements = driver.find_elements(By.CSS_SELECTOR,'.Price--priceWrapper--Q0Dn7pN')
        for price_element in price_elements:
            price_unit = price_element.find_element(By.CSS_SELECTOR,'.Price--unit--VNGKLAP').text
            price_int = price_element.find_element(By.CSS_SELECTOR,'.Price--priceInt--ZlsSi_M').text
            price_float = price_element.find_element(By.CSS_SELECTOR,'.Price--priceFloat--h2RR0RK').text
            realSales = price_element.find_element(By.CSS_SELECTOR,'.Price--realSales--FhTZc7U').text
            # Construct the price and add to the list
            price = f'{price_unit}{price_int}{price_float}'
            prices.append(price)
            Sold_num.append(realSales)

        # Extract shop info 
        shop_elements = driver.find_elements(By.CSS_SELECTOR,'.ShopInfo--shopName--rg6mGmy')
        Shop_name.extend([shop_element.text if shop_element.text else None for shop_element in shop_elements])
        Shop_link.extend([shop_element.get_attribute('href') if shop_element.text else None for shop_element in shop_elements])

        # Get link of product 
        product_elements = driver.find_elements(By.CSS_SELECTOR,'.Card--doubleCardWrapper--L2XFE73')
        Link_of_product.extend([product_element.get_attribute('href') for product_element in product_elements]) 
        time.sleep(5)
        
        next_button = driver.find_element(By.XPATH, "//span[text()='下一页']")
        next_button.click()
        time.sleep(2)
        
    homepage_button = driver.find_element(By.XPATH,"//span[text()='淘宝网首页']")
    homepage_button.click()
    time.sleep(2)
        
driver.quit()

#########################################################################################################################################################################################################################
# Create a DataFrame
df = pd.DataFrame({'Date': str(datetime.datetime.now().replace(microsecond=0)),'Title': title,'Prices': prices,'Img_urls': img_urls,'Sold_num': Sold_num,'Shop_name': Shop_name,'Shop_link': Shop_link,'Link_of_product': Link_of_product})
# Remove the first character (currency symbol) from 'Prices' column and convert to numeric
df['Prices'] = pd.to_numeric(df['Prices'].str.slice(1), errors='coerce')
df['Sold_num'] = df['Sold_num'].str.split('+').str[0]

# API file to ggsheet to translate then SAVE file again
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Authenticate with Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\\Users\Quynh Nhu\Downloads\Scrapping Data\vm-style-405901-b647d21322df.json", scope)
client = gspread.authorize(creds)

# Open the Google Sheet by title
google_sheet_title = 'API Test_VMStyle'
sheet = client.open(google_sheet_title).sheet1

# Append the new values to the sheet
sheet.append_rows(df.values.tolist())

print("Data successfully added to Google Sheets.")


# Define the columns and corresponding formulas
columns_formulas = [
    ('I', '=GOOGLETRANSLATE(B:B; "zh"; "vi")'),
    ('J', '=GOOGLETRANSLATE(E:E; "zh"; "vi")'),
    ('K', '=ROUND(C:C*GOOGLEFINANCE("CURRENCY:CNYVND"))'),
    ('L', '=GOOGLETRANSLATE(F:F; "zh"; "vi")')]


# Batch update formulas for each column
for column, formula in columns_formulas:
    column_index = ord(column) - ord('A') + 1
    # Get all values and indices from the sheet
    values_with_indices = [(i, value) for i, value in enumerate(sheet.col_values(1), start=1) if value != '']
    # Extract indices from the result
    non_empty_indices = len(values_with_indices)
    cell_range = f'{column}2:{column}{non_empty_indices}'  # Assuming your data starts from row 2
    # Build the batch update request
    request = {
        'repeatCell': {
            'range': {'sheetId': sheet.id, 'startRowIndex': 1, 'startColumnIndex': column_index - 1,
                      'endRowIndex': non_empty_indices , 'endColumnIndex': column_index},
            'cell': {'userEnteredValue': {'formulaValue': formula}},
            'fields': 'userEnteredValue.formulaValue'}
    }

    # Batch update request
    batch_update_body = {'requests': [request]}
    sheet.spreadsheet.batch_update(body=batch_update_body)

    
# SAVE data From ggsheet to file

# Get all values from the sheet
data = sheet.get_all_values()

# Convert the data to a Pandas DataFrame
df = pd.DataFrame(data[1:], columns=data[0])  # Assuming the first row contains column headers

# Create folder to save data:
key = list_keys    
folder_path = "C:\\Users\\TaoBao_Data" 
now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
filename = os.path.join(folder_path, f"{now}_{key}.xlsx")

# Save the DataFrame to an Excel file
excel_file_path = os.path.join(folder_path, filename)
df.to_excel(excel_file_path, index=False)

print(f'Data successfully saved to {excel_file_path}')
