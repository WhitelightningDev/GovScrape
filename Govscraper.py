import time
import logging
import pyodbc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
import re

# Set up logging
logging.basicConfig(filename="scraping.log", level=logging.INFO, format="%(asctime)s - %(message)s")

# SQL Server connection string
def get_db_connection():
    try:
        conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                              'SERVER=192.168.5.30;'
                              'DATABASE=Gov_Employees;'
                              'UID=dev;'
                              'PWD=Dev@2360')
        return conn
    except Exception as e:
        logging.error(f"Error connecting to the database: {e}")
        raise

# Function to update the database with the result
def update_employee_status(id_number, source, sector):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # Update the database with the new source and sector values
        update_query = """
            UPDATE Gov_Employees
            SET Source = ?, Sector = ?, Updated = 'true', InformationDate = GETDATE()
            WHERE ID_Number = ?
        """
        cursor.execute(update_query, (source, sector, id_number))
        conn.commit()  # Commit the transaction
        cursor.close()
        conn.close()

        logging.info(f"Database updated for ID {id_number}: Source = {source}, Sector = {sector}")
    except Exception as e:
        logging.error(f"Error updating the database for ID {id_number}: {e}")

# Function to initialize Selenium WebDriver without headless mode for debugging
def init_driver():
    try:
        options = Options()
        options.add_argument("--disable-gpu")
        options.add_argument("--start-maximized")  # Open browser maximized

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.get("https://www.dpsa.gov.za/resource_centre/psverification/")
        logging.info("Page loaded successfully.")
        return driver
    except WebDriverException as e:
        logging.error(f"Error initializing the WebDriver: {e}")
        raise

# Function to scrape the job information for an ID number
def get_job_info(driver, id_number):
    try:
        logging.info(f"Attempting to find the input field for ID: {id_number}")

        # Wait for the input field to be visible
        input_field = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))
        )
        logging.info("Input field found.")

        # Find the input field and enter the ID number
        input_field.clear()
        input_field.send_keys(id_number)

        # Wait for the submit button to be clickable and click it
        verify_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.ID, "Inputfield_submit"))
        )
        logging.info("Submit button found and clicked.")
        verify_button.click()

        # Wait for the response to appear (either successful or error response)
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".uk-text-center"))
        )
        logging.info("Response div found.")

        # Check if the response div has the 'uk-text-danger' class for "Not a Public Servant"
        try:
            response_div = driver.find_element(By.CSS_SELECTOR, ".uk-text-danger.uk-panel.uk-primary")
            logging.info(f"Response for ID {id_number}: Not a Public Servant")
            result_text = f"{id_number} - Not a Public Servant"
            source = "Not a Public Servant"
            sector = "None"  # Explicitly set sector to "None" for not a public servant
        except NoSuchElementException:
            # If not found, assume it's a valid response (i.e., successful)
            response_div = driver.find_element(By.CSS_SELECTOR, ".uk-text-success.uk-panel.uk-primary")
            result_text = response_div.text.strip()
            logging.info(f"Response for ID {id_number}: {result_text}")
            source = "Public Servant"

            # Extract province and sector
            if ":" in result_text:
                sector = result_text.split(":")[0].strip() + ":" + result_text.split(":")[1].strip()
            else:
                sector = result_text.strip()

            # Remove any digits from the sector and strip leading hyphens
            sector = re.sub(r'\d+', '', sector).strip()  # Remove digits
            sector = sector.lstrip('-').strip()  # Remove leading hyphen and extra spaces

        # Update the database with the result
        update_employee_status(id_number, source, sector)

        # Click "Do another search" button to reset the form
        do_another_search_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.uk-button.uk-button-primary.uk-grid-margin"))
        )
        do_another_search_button.click()
        logging.info("Form reset successfully.")

        # Wait for the input field to be visible again after reset
        WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))
        )

        return result_text

    except TimeoutException:
        logging.error(f"Timeout occurred while waiting for response for ID: {id_number}.")
        driver.refresh()
        return "Error: Timeout"

    except NoSuchElementException as e:
        logging.error(f"Element not found for ID {id_number}: {e}")
        return "Error: Element not found"

    except WebDriverException as e:
        logging.error(f"WebDriver error while processing ID {id_number}: {e}")
        return "Error: WebDriver"

    except Exception as e:
        logging.error(f"Unexpected error while processing ID {id_number}: {e}")
        return "Error: Unexpected error"

# Function to fetch all unprocessed IDs from the database
def get_all_ids_to_process():
    """Fetch all ID numbers where Updated = 'false'."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT ID_Number FROM Gov_Employees WHERE Updated = 'false' ORDER BY InformationDate ASC")
        rows = cursor.fetchall()

        return [row[0] for row in rows]  # Return list of ID numbers
    except Exception as e:
        logging.error(f"Error fetching ID numbers from the database: {e}")
        return []
    finally:
        conn.close()

# Function to extract processed IDs from the log file
def get_processed_ids_from_log():
    """
    Reads the 'scraping.log' file and extracts all IDs marked as processed.
    """
    processed_ids = set()
    try:
        with open("scraping.log", "r") as log_file:
            for line in log_file:
                match = re.search(r"Database updated for ID (\d+):", line)
                if match:
                    processed_ids.add(match.group(1))  # Extract the ID
    except FileNotFoundError:
        logging.warning("Log file not found. Proceeding without log check.")
    except Exception as e:
        logging.error(f"Error reading the log file: {e}")
    return processed_ids

# Main function to run the process
def process_ids():
    try:
        # Initialize Selenium WebDriver
        driver = init_driver()

        # Fetch all IDs to process from the database
        id_numbers = get_all_ids_to_process()

        # Fetch already processed IDs from the log file
        processed_ids = get_processed_ids_from_log()

        # Filter out IDs that are already in the log
        ids_to_process = [id_number for id_number in id_numbers if str(id_number) not in processed_ids]

        if ids_to_process:
            logging.info(f"IDs to process after filtering: {len(ids_to_process)}")
            for id_number in ids_to_process:
                logging.info(f"Processing ID: {id_number}")
                result = get_job_info(driver, id_number)

                if result == "Error: Timeout":
                    logging.warning(f"Skipping ID {id_number} due to timeout.")
                time.sleep(2)  # Optional delay to prevent overloading the server
        else:
            logging.info("No new IDs to process after filtering.")

        driver.quit()

    except Exception as e:
        logging.error(f"An error occurred during processing: {e}")

# Run the script
if __name__ == "__main__":
    process_ids()
