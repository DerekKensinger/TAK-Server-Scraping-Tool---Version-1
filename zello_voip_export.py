import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
import os

def main():
    # Get user input
    network_name = input("Enter network name: ")
    username = input("Enter username: ")
    password = input("Enter password: ")
    start_date = input("Enter start date and time (MM-DD-YYYY HH:MM): ")
    end_date = input("Enter end date and time (MM-DD-YYYY HH:MM): ")

    # Set up download directory
    download_dir = os.path.join(os.getcwd(), 'downloads')
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
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
        wait = WebDriverWait(driver, 20)  # Increased the timeout to 20 seconds

        # Step 1: Navigate to the initial sign-in page
        driver.get("https://zellowork.com/signin")

        # Step 2: Enter the network name
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.redirect-field.sign-in-to-zello-work__form__input")))
        network_field = driver.find_element(By.CSS_SELECTOR, "input.redirect-field.sign-in-to-zello-work__form__input")
        network_field.send_keys(network_name)

        # Step 3: Click the submit button
        submit_button = driver.find_element(By.CSS_SELECTOR, "button.btn.btn-xLarge.btn-primary.sign-in-to-zello-work__form__button")
        submit_button.click()

        # Wait for the login page to load
        wait.until(EC.element_to_be_clickable((By.NAME, 'username')))

        # Step 4: Enter username and password
        username_field = driver.find_element(By.NAME, 'username')
        username_field.send_keys(username)

        password_field = driver.find_element(By.NAME, 'password')
        password_field.send_keys(password)

        # Step 5: Click the login button
        login_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
        login_button.click()

        # Wait for login to complete by checking the URL or a specific element on the next page
        wait.until(EC.url_contains(f"{network_name}.zellowork.com"))

        # Alternatively, wait for a specific element that appears after login
        # wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'selector_of_element_after_login')))

        # Step 6: Navigate to history page
        history_url = f"https://{network_name}.zellowork.com/history"
        driver.get(history_url)

        # Wait for the page to load
        wait.until(EC.element_to_be_clickable((By.NAME, 'start-date')))

        # Set timeframe using JavaScript
        set_date_with_js(driver, 'start-date', start_date)
        set_date_with_js(driver, 'end-date', end_date)

        # Wait for the page to update the data based on new dates
        time.sleep(5)

        # Click the Export Metadata button
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[fname="export-metadata"]')))
        export_button = driver.find_element(By.CSS_SELECTOR, 'button[fname="export-metadata"]')
        export_button.click()

        # Handle the modal window if necessary
        handle_export_modal(driver, wait)

        # Wait for download to complete
        wait_for_downloads(download_dir)

        # Rename the downloaded file
        rename_downloaded_file(download_dir)

    except Exception as e:
        print(f"An error occurred: {e}")
        print("Current URL:", driver.current_url)
        driver.save_screenshot('error_screenshot.png')
    finally:
        # Close the driver
        driver.quit()

def set_date_with_js(driver, field_name, date_value):
    # Locate the date field
    date_field = driver.find_element(By.NAME, field_name)
    # Use JavaScript to set the date value
    driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", date_field, date_value)

def handle_export_modal(driver, wait):
    try:
        # Wait for the modal to appear
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.modal-download-selected.history-metadata')))

        # If there's a confirm button inside the modal, click it
        confirm_button = driver.find_element(By.CSS_SELECTOR, 'button[fname="download-metadata"]')
        confirm_button.click()
    except:
        # If no modal appears, proceed
        pass

def wait_for_downloads(download_dir):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 60:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(download_dir)
        for fname in files:
            if fname.endswith('.crdownload') or fname.endswith('.part'):
                dl_wait = True
        seconds += 1

def rename_downloaded_file(download_dir):
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
