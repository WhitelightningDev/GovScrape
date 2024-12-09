import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
import logging

# Set up logging
logging.basicConfig(filename="scraping.log", level=logging.INFO, format="%(asctime)s - %(message)s")

# Function to initialize Selenium WebDriver without headless mode for debugging
def init_driver():
    try:
        options = Options()
        # Disable headless mode for debugging
        # options.add_argument("--headless")  # Uncomment to enable headless mode for production
        options.add_argument("--disable-gpu")
        options.add_argument("--start-maximized")  # Open browser maximized

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.get("https://www.dpsa.gov.za/resource_centre/psverification/")
        logging.info("Page loaded successfully.")
        return driver
    except WebDriverException as e:
        logging.error(f"Error initializing the WebDriver: {e}")
        raise

# Function to monitor changes in the container
def monitor_container_change(driver, container_selector, previous_content=None):
    try:
        container = driver.find_element(By.CSS_SELECTOR, container_selector)
        current_content = container.get_attribute("outerHTML")

        if previous_content and current_content != previous_content:
            logging.info("Content has changed!")
            return current_content
        elif not previous_content:
            return current_content
        else:
            logging.info("No changes detected in the container.")
            return previous_content
    except NoSuchElementException:
        logging.error("Container not found!")
        return previous_content
    except Exception as e:
        logging.error(f"Error monitoring container change: {e}")
        return previous_content

# Function to scrape the job information for an ID number
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
        except NoSuchElementException:
            # If not found, assume it's a valid response (i.e., successful)
            response_div = driver.find_element(By.CSS_SELECTOR, ".uk-text-success.uk-panel.uk-primary")
            result_text = response_div.text.strip()
            logging.info(f"Response for ID {id_number}: {result_text}")

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

# Function to process the ID numbers and save the result to an Excel file
def process_ids(input_file, output_file):
    try:
        # Read ID numbers from Excel file
        df = pd.read_excel(input_file)
        id_numbers = df['ID Number'].tolist()

        # Initialize Selenium driver
        driver = init_driver()

        # Create a list to hold the results
        results = []

        # Loop through the ID numbers and scrape the data
        for id_number in id_numbers:
            logging.info(f"Processing ID: {id_number}")
            result = get_job_info(driver, id_number)
            results.append({"ID Number": id_number, "Result": result})

            if result == "Error: Timeout":
                logging.warning(f"Skipping ID {id_number} due to timeout.")
            time.sleep(2)

        # Save results to an Excel file
        results_df = pd.DataFrame(results)
        results_df.to_excel(output_file, index=False)
        logging.info(f"Results saved to {output_file}")

        driver.quit()

    except FileNotFoundError as e:
        logging.error(f"Error: The input file was not found. Error: {e}")
    except PermissionError as e:
        logging.error(f"Error: Permission denied while accessing the file. Error: {e}")
    except Exception as e:
        logging.error(f"Unexpected error while processing the file: {e}")

# Main function
if __name__ == "__main__":
    input_file = "idnumber.xlsx"  # Replace with your input Excel file
    output_file = "output_jobs.xlsx"  # Output file to save the results

    try:
        process_ids(input_file, output_file)
        logging.info("Scraping completed.")
    except Exception as e:
        logging.error(f"An error occurred during scraping: {e}")
