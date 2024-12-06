import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException


# Function to initialize Selenium WebDriver with headless mode
def init_driver():
    try:
        options = Options()
        options.add_argument("--headless")  # Run in headless mode (no GUI)
        options.add_argument("--disable-gpu")  # Disable GPU for smoother running

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.get("https://www.dpsa.gov.za/resource_centre/psverification/")
        return driver
    except WebDriverException as e:
        print(f"Error initializing the WebDriver: {e}")
        raise


# Function to scrape the job information for an ID number
def get_job_info(driver, id_number):
    try:
        # Wait for the input field to be visible
        input_field = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))  # Wait for input field to appear
        )

        # Find the input field and enter the ID number
        input_field.clear()
        input_field.send_keys(id_number)

        # Find and click the submit button (Verify)
        verify_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "Inputfield_submit"))  # Button ID for Submit
        )
        verify_button.click()

        # Wait for the response to appear in the form of a state change in the HTML
        WebDriverWait(driver, 15).until(
            EC.text_to_be_present_in_element((By.CLASS_NAME, "uk-text-danger"), "Not a Public Servant")
        )

        # Once the response is detected, refresh the page
        driver.refresh()

        # Wait for the page to reload completely before the next ID can be processed
        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))  # Ensure input field is ready again
        )

        print(f"Processed ID: {id_number}")
        return "Processed"
    except (NoSuchElementException, TimeoutException) as e:
        print(f"Error while processing ID {id_number}: Element not found or timeout occurred. Error: {e}")
        return "Error"
    except WebDriverException as e:
        print(f"WebDriver error while processing ID {id_number}: {e}")
        return "Error"
    except Exception as e:
        print(f"Unexpected error while processing ID {id_number}: {e}")
        return "Error"


# Function to process the ID numbers and save the result to an Excel file
def process_ids(input_file, output_file):
    try:
        # Read ID numbers from Excel file
        df = pd.read_excel(input_file)
        id_numbers = df['ID Number'].tolist()

        # Initialize Selenium driver
        driver = init_driver()

        # Loop through the ID numbers and scrape the data
        for id_number in id_numbers:
            print(f"Processing ID: {id_number}")
            result = get_job_info(driver, id_number)
            if result == "Error":
                print(f"Skipping ID: {id_number} due to error.")
            time.sleep(2)  # Adjust the sleep time to allow enough time between searches

        driver.quit()
    except FileNotFoundError as e:
        print(f"Error: The input file was not found. Please check the file path. Error: {e}")
    except PermissionError as e:
        print(f"Error: Permission denied while accessing the file. Error: {e}")
    except Exception as e:
        print(f"Unexpected error while processing the file: {e}")


# Main function
if __name__ == "__main__":
    input_file = "idnumber.xlsx"  # Replace with your input Excel file
    output_file = "output_jobs.xlsx"  # Output file to save the results

    try:
        process_ids(input_file, output_file)
        print("Scraping completed.")
    except Exception as e:
        print(f"An error occurred during scraping: {e}")
