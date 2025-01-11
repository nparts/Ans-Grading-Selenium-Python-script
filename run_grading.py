import os
import re
import subprocess
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.file_detector import LocalFileDetector
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

# Define assignment url (make sure you go to the url tab in the browser and copy the url)
url="https://ans.app/assignments/1229647/results"

# Load the spreadsheet (You can export the spreadsheet as nl excel file from ANS)
file_path = '/Users/nielsarts/Library/CloudStorage/OneDrive-WindesheimOffice365/Beoordelings formulier HW DP5 AB.xlsx'
df = pd.read_excel(file_path)

# Define the mapper for values
def map_value(value):
    mapping = {
        -1: "[1]",     # Niet aanwezig
        0: "[2]",   # In ontwikkeling
        1: "[3]",   # Op niveau
        2: "[4]"    # Boven niveau
        }
    return mapping.get(value, None)

# Define the mapper for columns (This can be added based on the columns in the spreadsheet)
def map_column(column):
    column_mapping = {
        '1a': "[1]",
        '1b': "[2]",
        '1c': "[3]",
        '1d': "[4]",
        '1e': "[5]",
        '1f': "[6]"
    }
    return column_mapping.get(column, None)

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

    # Loop through each row from the spreadsheet
    for index, row in df.iterrows():
        student_number = row['Studentnummer']

        # Open the assignment url
        driver.get(url)
        # Fill in the student number in the search field and click on the result
        driver.find_element(By.ID, "search_items").click()
        driver.find_element(By.ID, "search_items").send_keys(str(student_number))
        WebDriverWait(driver, 30).until(expected_conditions.text_to_be_present_in_element((By.XPATH, "/html/body/div[5]/div/main/div[6]/div/table/tbody/tr/td[5]"), str(student_number)))
        driver.find_element(By.XPATH, "/html/body/div[5]/div/main/div[6]/div/table/tbody/tr/td[2]/div/a").click()
        # Open the student assignment for grading
        driver.implicitly_wait(1) #seconds
        WebDriverWait(driver, 30).until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, ".font-heading > .d-flex")))
        driver.find_element(By.CSS_SELECTOR, ".font-heading > .d-flex").click()
        driver.implicitly_wait(2) #seconds

        # Loop through each column and click on the value
        for col in ['1a', '1b', '1c', '1d']:
            driver.implicitly_wait(2)
            value = row[col]
            # Map the value and column to the correct xpath
            mapped_value = map_value(value)
            mapped_column = map_column(col)
            if mapped_value is not None and mapped_column is not None:
                xpath = f"/html/body/main/div[1]/section/div[3]/div[1]/div[2]/div{mapped_column}/ul[1]/li{mapped_value}/div[1]/a"
                # xpath = f"//body[@id='body']/main/div[1]/section/div[3]/div[1]/div[2]/div{mapped_column}/ul[2]/li{mapped_value}/div[0]/a"
                      #  //*[@id="body"]/main/div[1]/section/div[3]/div[1]/div[2]/div[1]            /ul[1]/li[4]           /div[1]/a
                time.sleep(1)
                driver.find_element(By.XPATH, xpath).click()
                driver.implicitly_wait(2)
                # print which student number and column is clicked
                print(f"Student number: {student_number}, Column: {col}, Value: {value}")
        driver.implicitly_wait(2)
    driver.quit()
