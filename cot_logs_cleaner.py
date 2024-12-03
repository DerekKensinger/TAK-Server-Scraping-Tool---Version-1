from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import os
import zipfile
import shutil
import tempfile

def main():
    # Get user input
    network_name = input("Enter network name: ")
    start_date = input("Enter start date and time (MM-DD-YYYY HH:MM): ")
    end_date = input("Enter end date and time (MM-DD-YYYY HH:MM): ")

    # Set up download directory
    download_dir = os.path.join(os.getcwd(), 'downloads')
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

    # Create a temporary user data directory
    temp_user_data_dir = tempfile.mkdtemp()

    # Copy your automation profile to the temporary directory
    original_profile_path = r'C:\Users\dkens\AppData\Local\Google\Chrome\User Data\Default'  
    temp_profile_path = os.path.join(temp_user_data_dir, 'Default')
    shutil.copytree(original_profile_path, temp_profile_path)

    chrome_options.add_argument(f"--user-data-dir={temp_user_data_dir}")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # Initialize the driver (update the path to your chromedriver if necessary)
    driver = webdriver.Chrome(options=chrome_options)

    try:
        # Navigate to history page directly
        history_url = f"https://{network_name}.zellowork.com/history"
        driver.get(history_url)

        # Wait for the page to load
        time.sleep(5)  # Adjust as necessary

        # Check if authentication is required
        if "login" in driver.current_url.lower() or "signin" in driver.current_url.lower():
            print("Authentication is required. Please log in manually.")
            input("Press Enter after you have logged in...")
            # Wait for the page to load after login
            time.sleep(5)

        # Verify we're on the history page
        print(f"Current URL: {driver.current_url}")
        driver.save_screenshot('page_before_setting_dates.png')

        # Handle iframe if present
        handle_iframe(driver)

        # Wait for the start-date element to be present
        wait = WebDriverWait(driver, 20)  # 20 seconds timeout
        try:
            start_date_element = wait.until(EC.presence_of_element_located((By.NAME, 'start-date')))
        except Exception as e:
            print(f"Failed to find the start-date element: {e}")
            driver.save_screenshot('start_date_not_found.png')
            return  # Exit the script or handle the error as appropriate

        # Proceed to set the date range
        set_date_with_js(driver, 'start-date', start_date)
        set_date_with_js(driver, 'end-date', end_date)

        # Wait for the page to update the data based on new dates
        time.sleep(5)

        # Click the Export Metadata button
        export_button = driver.find_element(By.CSS_SELECTOR, 'button[fname="export-metadata"]')
        export_button.click()

        # Handle the modal window if necessary
        handle_export_modal(driver)

        # Wait for download to complete
        wait_for_downloads(download_dir)

        # Rename and extract the downloaded file
        rename_and_extract_zip(download_dir)

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the driver
        driver.quit()
        # Remove the temporary user data directory
        shutil.rmtree(temp_user_data_dir)

def handle_iframe(driver):
    # Check if there are any iframes on the page
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    print(f"Found {len(iframes)} iframe(s) on the page.")

    if len(iframes) > 0:
        # If there's more than one iframe, you may need to adjust this
        driver.switch_to.frame(iframes[0])
        print("Switched to iframe.")

def set_date_with_js(driver, field_name, date_value):
    # Locate the date field
    date_field = driver.find_element(By.NAME, field_name)
    # Use JavaScript to set the date value
    driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", date_field, date_value)

def handle_export_modal(driver):
    try:
        # Wait for the modal to appear
        time.sleep(2)  # Adjust as necessary
        confirm_button = driver.find_element(By.CSS_SELECTOR, 'button[fname="download-metadata"]')
        confirm_button.click()
    except:
        # If no modal appears, proceed
        pass

def wait_for_downloads(download_dir):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 120:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(download_dir)
        for fname in files:
            if fname.endswith('.crdownload') or fname.endswith('.part'):
                dl_wait = True
        seconds += 1

def rename_and_extract_zip(download_dir):
    files = os.listdir(download_dir)
    zip_files = [f for f in files if f.endswith('.zip')]

    if len(zip_files) == 1:
        original_zip_file = os.path.join(download_dir, zip_files[0])
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        new_csv_filename = f"Zello VOIP Export {timestamp}.csv"
        extract_and_rename_csv(original_zip_file, download_dir, new_csv_filename)
        print(f"File saved as {os.path.join(download_dir, new_csv_filename)}")
    else:
        print("Error: Expected one zip file in download directory.")

def extract_and_rename_csv(zip_file_path, extract_to_dir, new_csv_filename):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # Assuming there is only one CSV file in the zip
        csv_files_in_zip = [f for f in zip_ref.namelist() if f.endswith('.csv')]
        if len(csv_files_in_zip) == 1:
            csv_filename_in_zip = csv_files_in_zip[0]
            zip_ref.extract(csv_filename_in_zip, extract_to_dir)
            original_csv_path = os.path.join(extract_to_dir, csv_filename_in_zip)
            new_csv_path = os.path.join(extract_to_dir, new_csv_filename)
            os.rename(original_csv_path, new_csv_path)
            # Optionally, remove the original zip file
            os.remove(zip_file_path)
        else:
            print("Error: Expected one CSV file in the zip archive.")

if __name__ == "__main__":
    main()
