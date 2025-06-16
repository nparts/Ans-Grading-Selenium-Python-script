from time import sleep
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait

# Define assignment url (make sure you go to the url tab in the browser and copy the url)
# DP31
url="https://ans.app/assignments/1423365/results"

# Load the spreadsheet (You can export the spreadsheet as nl excel file from ANS)
file_path = '/Users/nielsarts/Desktop/Beoordelingsformulier.xlsx'
df = pd.read_excel(file_path, converters={
    "Studentnummer":str,
    '14a':str,
    '14a-opmerking':str,
    '14b':str,
    '14c':str,
    '14d':str,
    '31a':str,
    '31a-opmerknig':str,
    '31b':str,
    '31c':str,
}, sheet_name='CD')

# Define the mapper for values
def map_value(value):
    mapping = {
        '-1': "[1]",     # Niet aanwezig
        '0': "[2]",   # In ontwikkeling
        '1': "[3]",   # Op niveau
        '2': "[4]"    # Boven niveau
        }
    return mapping.get(value, None)

# Define the mapper for columns (This can be added based on the columns in the spreadsheet)
def map_column(column):
    column_mapping = {
        '31a': "[1]",
        '31b': "[2]",
        '31c': "[3]",
    }
    return column_mapping.get(column, None)

# Define the student number to skip until (set to None to process all rows)
SKIP_TILL_INCLUDING_STUDENT_NUMBER = None # Replace with the desired student number or None

def test_process_results():
    """
    Automates the process of grading student assignments using Selenium WebDriver.
    This function initializes a Selenium WebDriver with specified Chrome options, 
    iterates through each row in a DataFrame, and performs the following actions:
    - Opens the assignment URL.
    - Searches for a student number.
    - Opens the student's assignment for grading.
    - Clicks on specific values in the assignment based on the DataFrame columns.
    Prerequisites:
    - Ensure that Selenium WebDriver is installed and the driver path is set correctly.
    - Chrome driver should be running on port 9222.
    - User should be logged in manually in ANS before running the script.
    Args:
        None
    Returns:
        None
    Raises:
        Various Selenium exceptions if elements are not found or actions cannot be performed.
    """

    # Initialize Selenium WebDriver (Make sure you have installed it and set the driver path correctly)
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir=/Users/nielsarts/ChromeProfile")
    # make sure your chrome driver is running on port 9222 (first login manually in ANS before running the script)
    options.debugger_address="localhost:9222"
    driver = webdriver.Chrome(options=options)

    # Flag to determine when to start processing
    start_processing = SKIP_TILL_INCLUDING_STUDENT_NUMBER is None

    # Loop through each row from the spreadsheet
    for index, row in df.iterrows():
        student_number = row['Studentnummer']

        # Check if we should start processing
        if not start_processing:
            if student_number == SKIP_TILL_INCLUDING_STUDENT_NUMBER:
                start_processing = True
                continue
            else:
                print(f"Skipping student number: {student_number}")
                continue

        print(f"Grading student number: {student_number}")
        # Open the assignment url
        driver.get(url)
        # Fill in the student number in the search field and click on the result
        search = driver.find_element(By.ID, "search_items")
        search.click()
        search.send_keys(student_number)

        WebDriverWait(driver, 5).until(expected_conditions.text_to_be_present_in_element((By.XPATH, "/html/body/div[5]/div/main/div[4]/div/table/tbody/tr/td[5]"), str(student_number)))
        assignment = driver.find_element(By.XPATH, "/html/body/div[5]/div/main/div[4]/div/table/tbody/tr/td[2]/div/a")
        assignment.click()

        # Open the student assignment for grading
        WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/main/div/div[1]/div/div/a[2]")))
        bekijk = driver.find_element(By.XPATH, "/html/body/div[4]/div/main/div/div[1]/div/div/a[2]")
        bekijk.click()
        
        # Loop through each column and click on the value
        for col in ['31a','31b', '31c']:  # Add more columns if needed
            value = row[col]
            # Map the value and column to the correct xpath
            mapped_value = map_value(value)
            mapped_column = map_column(col)
            comment = row.get(f"{col}-opmerking", "")
        
            if mapped_value is not None and mapped_column is not None:
                        # /html/body/main/div[1]/section/div[3]/div/div[1]/div[1]            /ul[2]/li[3]           /div[1]/a
                xpath = f"/html/body/main/div[1]/section/div[3]/div/div[1]/div{mapped_column}/ul[2]/li{mapped_value}/div[1]/a"
                WebDriverWait(driver, 15).until(expected_conditions.element_to_be_clickable((By.XPATH, xpath)))
                beoordeel = driver.find_element(By.XPATH, xpath)
                beoordeel.click()
                print(f"Student number: {student_number}, Column: {col}, Value: {value}")
                
                # If there is a comment and is not empty, fill it in
                if (comment):
                    if( not isinstance(comment, str)) or (not comment.strip()):
                        continue

                    print(f"Filling in comment for {col}: {comment}")
                                           # /html/body/main/div[1]/section/div[3]/div/div[1]/div[1]            /ul[2]/li[3]           /div[2]/div[3]/div/a
                    comment_button_xpath = f"/html/body/main/div[1]/section/div[3]/div/div[1]/div{mapped_column}/ul[2]/li{mapped_value}/div[2]/div[3]/div/a"
                    WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((By.XPATH, comment_button_xpath)))
                    comment_button = driver.find_element(By.XPATH, comment_button_xpath)
                    comment_button.click()

                    # /html/body/main/div[1]/section/div[3]/div/div[1]/div[1]/ul[2]/li[3]/div[2]/div[3]/div[2]/form/div[1]/div/div[2]
                    comment_input_xpath = f"/html/body/main/div[1]/section/div[3]/div/div[1]/div{mapped_column}/ul[2]/li{mapped_value}/div[2]/div[3]/div[2]/form/div[1]/div/div[2]"
                    # wait till input is clickable
                    WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((By.XPATH, comment_input_xpath)))
                    comment_input = driver.find_element(By.XPATH, comment_input_xpath)
                    comment_input.click()
                    comment_input.send_keys(comment)
                    # wait for the keys to be sent
                    sleep(1)

                    # /html/body/main/div[1]/section/div[3]/div/div[1]/div[1]/ul[2]/li[3]/div[2]/div[3]/div[2]/form/div[2]/button[2]
                    save_button_xpath = f"/html/body/main/div[1]/section/div[3]/div/div[1]/div{mapped_column}/ul[2]/li{mapped_value}/div[2]/div[3]/div[2]/form/div[2]/button[2]"
                    # Wait for the save button to be clickable element is a button with type submit within a form
                    WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((By.XPATH, save_button_xpath)))
                    save_button = driver.find_element(By.XPATH, save_button_xpath)
                    save_button.click()
                    print(f"Comment saved for {col}: {comment}")
            else:
                print(f"Invalid value or column: {value}, {col}")
    driver.quit()
