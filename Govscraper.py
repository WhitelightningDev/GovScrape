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

# Function to initialize Selenium WebDriver with headless mode
def init_driver():
    try:
        options = Options()
        options.add_argument("--headless")  # Run in headless mode (no GUI)
        options.add_argument("--disable-gpu")  # Disable GPU for smoother running

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.get("https://www.dpsa.gov.za/resource_centre/psverification/")
        print("Page loaded successfully.")
        return driver
    except WebDriverException as e:
        print(f"Error initializing the WebDriver: {e}")
        raise


# Function to scrape the job information for an ID number
def get_job_info(driver, id_number):
    try:
        print(f"Attempting to find the input field for ID: {id_number}")

        # Wait for the input field to be visible with 1-minute timeout
        input_field = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))  # Wait for input field to appear
        )

        print("Input field found.")

        # Find the input field and enter the ID number
        input_field.clear()
        input_field.send_keys(id_number)

        # Find and click the submit button (Verify) with 1-minute timeout
        verify_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.ID, "Inputfield_submit"))  # Button ID for Submit
        )
        print("Submit button found and clicked.")
        verify_button.click()

        # Wait for the response to appear with 1-minute timeout
        try:
            print("Waiting for response...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, ".uk-text-danger.uk-panel.uk-primary")
                )  # Wait for the response div containing the text
            )
            # Check if the text matches the expected response
            response_element = driver.find_element(By.CSS_SELECTOR, ".uk-text-danger.uk-panel.uk-primary")
            if "Not a Public Servant" in response_element.text:
                print(f"Response for ID {id_number}: {response_element.text}")
                return response_element.text
            else:
                print(f"Unexpected response for ID {id_number}: {response_element.text}")
                return response_element.text

        except TimeoutException:
            print(f"Timeout occurred while waiting for response for ID: {id_number}. Attempting a page refresh.")
            driver.refresh()  # Refresh page if timeout occurs

        # Check for any URL or DOM change after clicking the submit button
        current_url = driver.current_url
        print(f"Current URL after submitting ID: {current_url}")
        time.sleep(5)  # Wait for the page to fully load the response

        # Monitor for state changes by checking the DOM for relevant elements
        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".uk-text-danger.uk-panel.uk-primary"))
            )
            print("State change detected: Response displayed.")
        except TimeoutException:
            print(f"Timeout occurred while waiting for state change after ID {id_number}. Checking for issues.")

        # Refresh the page after processing the current ID to reset the input field for the next ID
        print(f"Refreshing the page after processing ID {id_number}.")
        driver.refresh()

        # Wait for the page to reload completely before the next ID can be processed
        WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))  # Ensure input field is ready again
        )
        print("Page refreshed and ready for next ID.")
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

        # Create a list to hold the results
        results = []

        # Loop through the ID numbers and scrape the data
        for id_number in id_numbers:
            print(f"Processing ID: {id_number}")
            result = get_job_info(driver, id_number)
            results.append({"ID Number": id_number, "Result": result})

            if result == "Error":
                print(f"Skipping ID: {id_number} due to error.")
            time.sleep(2)  # Adjust the sleep time to allow enough time between searches

        # Save results to an Excel file
        results_df = pd.DataFrame(results)
        results_df.to_excel(output_file, index=False)
        print(f"Results saved to {output_file}")

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
