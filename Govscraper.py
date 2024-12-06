import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.frame import Frame
from webdriver_manager.chrome import ChromeDriverManager


# Function to initialize Selenium WebDriver with headless mode
def init_driver():
    options = Options()
    options.add_argument("--headless")  # Run in headless mode (no GUI)
    options.add_argument("--disable-gpu")  # Disable GPU for smoother running

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://www.dpsa.gov.za/resource_centre/psverification/")
    return driver


# Function to scrape the job information for an ID number
def get_job_info(driver, id_number):
    try:
        # Wait for the input field to be visible
        input_field = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "idNumber"))  # Wait for input field to appear
        )

        # Check if the input field is within an iframe
        if driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))

        # Find the input field and enter the ID number
        input_field.clear()
        input_field.send_keys(id_number)

        # Find and click the verify button
        verify_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Verify')]"))
        )
        verify_button.click()

        # Wait for the result to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "uk-text-danger"))
        )

        # Extract result text
        result_div = driver.find_element(By.CLASS_NAME, "uk-text-danger")
        job_description = result_div.text.strip()

        # Check if the "DO ANOTHER SEARCH" button is displayed after the result
        try:
            do_another_button = driver.find_element(By.XPATH, "//button[contains(text(), 'DO ANOTHER SEARCH')]")
            if do_another_button.is_displayed():
                # Click it to clear the form for the next search
                do_another_button.click()
                time.sleep(1)  # Wait for the form to reset
        except Exception:
            pass

        return job_description
    except Exception as e:
        print(f"Error for ID {id_number}: {e}")
        return "Error"


# Function to process the ID numbers and save the result to an Excel file
def process_ids(input_file, output_file):
    # Read ID numbers from Excel file
    df = pd.read_excel(input_file)
    id_numbers = df['ID Number'].tolist()

    # Initialize Selenium driver
    driver = init_driver()

    # Prepare a list to store results
    results = []

    # Loop through the ID numbers and scrape the data
    for id_number in id_numbers:
        print(f"Processing ID: {id_number}")
        job_description = get_job_info(driver, id_number)
        results.append({
            "ID Number": id_number,
            "Job Description": job_description
        })

    # Convert results to DataFrame and save to Excel
    try:
        results_df = pd.DataFrame(results)
        results_df.to_excel(output_file, index=False)
        print(f"Data saved to {output_file}")
    except PermissionError:
        print(f"Permission denied: Ensure the file '{output_file}' is not open in any other application.")

    driver.quit()


# Main function
if __name__ == "__main__":
    input_file = "idnumber.xlsx"  # Replace with your input Excel file
    output_file = "output_jobs.xlsx"  # Output file to save the results

    process_ids(input_file, output_file)
    print("Scraping completed and data saved.")
