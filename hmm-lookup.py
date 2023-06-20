import undetected_chromedriver as uc
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook
from datetime import datetime

def wait_for_content(driver, element):
    # Wait for the JavaScript to fill in elements
    wait = WebDriverWait(driver, 10)  # Maximum wait time of 10 seconds
    element_locator = (By.XPATH, element)
    wait.until(EC.presence_of_element_located(element_locator))

def fill_input_initial(driver, tracker):
    wait_for_content(driver, "//input[@name='srchBkgNo1']")
    input_box = driver.find_element(By.XPATH, "//input[@name='srchBkgNo1']")
    input_box.send_keys(tracker)
    
    wait_for_content(driver, "//button[normalize-space()='Retrieve']")
    retrieve_button = driver.find_element(By.XPATH, "//button[normalize-space()='Retrieve']")
    retrieve_button.click()

def fill_input_sub(driver, tracker):
    wait_for_content(driver, "//input[@id='esvcGlobalQuery']")
    input_box = driver.find_element(By.XPATH, "//input[@id='esvcGlobalQuery']")
    input_box.send_keys(tracker)
    
def retrieve_date_info(driver):
    try:
        wait_for_content(driver, ".//div[@id='cntrChangeArea']//table//tbody/tr[3]")
        date_element_row = driver.find_element(By.XPATH, ".//div[@id='cntrChangeArea']//table//tbody/tr[3]")
        date_element_last_td = date_element_row.find_elements(By.XPATH, "./td")
        
        date_unformatted = date_element_last_td[-1].text
        date_formatted = date_unformatted.split(" ", 1)
        print(date_formatted[0])
        
        return date_formatted[0]
    
    except Exception as e:
        print(f"Could not find date on page. {e} \n")
        return 'N/A'
    
def format_date(date):
    # Parse the input string into a datetime object
    date_object = datetime.strptime(date, "%Y-%m-%d")

    # Format the date as "month/day"
    formatted_date = date_object.strftime("%m/%d")
    return formatted_date
    
# Setup excel workbook
workbook = Workbook()
worksheet = workbook.active
worksheet.title = "Shipping Date Changes"
worksheet.column_dimensions['A'].width = 25

# Create a new instance of the Firefox driver
driver = uc.Chrome(use_subprocess=True)
driver.get('https://www.hmm21.com/e-service/general/trackNTrace/TrackNTrace.do?&blNo=')

# Get list of MSC tracking numbers
list_tracking_numbers = open("list-trackers.txt", "r").readlines()

for entry in list_tracking_numbers:
    if entry == list_tracking_numbers[0]:
        # First search using different webpage
        fill_input_initial(driver, entry)
    else:
        # Subsequent searches uses top search bar
        fill_input_sub(driver, entry)
            
    date = retrieve_date_info(driver)
    
    try:
        date = format_date(date)
    except ValueError:
        date = "No ETA Found"
    row = [entry.strip(), date]
    
    # append row into worksheet
    worksheet.append(row)
    

workbook.save("output/hmm_shipping_dates_changes.xlsx")

driver.quit()