from time import sleep
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait

# Define assignment url (make sure you go to the url tab in the browser and copy the url)
# DP17
# url="https://ans.app/assignments/1376037/results"
# DP18
# url="https://ans.app/assignments/1376065/results"
# DP20
# url="https://ans.app/assignments/1376077/results"
# DP21
url="https://ans.app/assignments/1376091/results"

# Load the spreadsheet (You can export the spreadsheet as nl excel file from ANS)
file_path = '/Users/nielsarts/Desktop/Beoordelings formulier - Sprint 2.xlsx'
df = pd.read_excel(file_path, converters={
    "Studentnummer":str,
    '17a':str,
    '17b':str,
    '17c':str,
    '17d':str,
    '18a':str,
    '18b':str,
    '20a':str,
    '20b':str,
    '21a':str,
    '21b':str,
    '21c':str,
    '21d':str
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
        '21a': "[1]",
        '21b': "[2]",
        '21c': "[3]",
        '21d': "[4]",
    }
    return column_mapping.get(column, None)

# Define the student number to skip until (set to None to process all rows)
SKIP_TILL_INCLUDING_STUDENT_NUMBER = "1210564"  # Replace with the desired student number or None

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
        for col in ['21a','21b','21c','21d']:  # Add more columns if needed
            value = row[col]
            # Map the value and column to the correct xpath
            mapped_value = map_value(value)
            mapped_column = map_column(col)
            if mapped_value is not None and mapped_column is not None:
                xpath = f"/html/body/main/div[1]/section/div[3]/div/div[1]/div{mapped_column}/ul[2]/li{mapped_value}/div[1]/a"
                WebDriverWait(driver, 5).until(expected_conditions.element_to_be_clickable((By.XPATH, xpath)))
                beoordeel = driver.find_element(By.XPATH, xpath)
                beoordeel.click()
                # print which student number and column is clicked
                print(f"Student number: {student_number}, Column: {col}, Value: {value}")
            else:
                print(f"Invalid value or column: {value}, {col}")
    driver.quit()
