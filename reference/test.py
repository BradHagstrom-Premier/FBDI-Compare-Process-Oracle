### Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import StaleElementReferenceException
import os
import shutil
import requests
import time
import xlwings as xw
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def create_folders_and_clear():
    # List of folder names to create or clear
    folders_to_create_or_clear = ['Blank Copies', 'Originals']

    # Current directory
    current_directory = os.getcwd()

    # Loop through the list and create each folder if it doesn't exist, or clear it if it does
    for folder in folders_to_create_or_clear:
        folder_path = os.path.join(current_directory, folder)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Created folder: {folder_path}")
        else:
            # If the folder exists, delete its contents
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                    print(f"Deleted {file_path}")
                except Exception as e:
                    print(f"Failed to delete {file_path}. Reason: {e}")

def download_with_retry(session, url, timeout=(5, 15), retries=5, backoff_factor=1):
    retry = Retry(
        total=retries,
        read=retries,
        connect=retries,
        backoff_factor=backoff_factor,
        status_forcelist=(500, 502, 504),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    response = session.get(url, stream=True, timeout=timeout)
    response.raise_for_status()
    return response


# def download_files(driver, download_path):
#     base_urls = [
#         'https://docs.oracle.com/en/cloud/saas/project-management/25b/oefpp/index.html',
#         'https://docs.oracle.com/en/cloud/saas/financials/25b/oefbf/index.html',
#         'https://docs.oracle.com/en/cloud/saas/procurement/25b/oefbp/index.html',
#         'https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/25b/oefsc/index.html'
#     ]
#     for base_url in base_urls:
#         driver.get(base_url)
#         WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'navigationDrawer')))

#         category_elements = driver.find_elements(By.XPATH, '//*[@id="navigationDrawer"]//a')
#         category_urls = [element.get_attribute('href') for element in category_elements]
#         print(category_urls)

#         session = requests.Session()
#         for category_url in category_urls:
#             try:
#                 driver.get(category_url)
                
#                 WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href$='.xlsm']")))

#                 attempt = 0
#                 max_attempts = 3
#                 while attempt < max_attempts:
#                     try:
#                         download_links = driver.find_elements(By.CSS_SELECTOR, "a[href$='.xlsm']")
#                         for link in download_links:
#                             download_url = link.get_attribute('href')
#                             if download_url.endswith('.xlsm'):
#                                 local_filename = download_url.split('/')[-1]
#                                 try:
#                                     response = download_with_retry(session, download_url)
#                                     with open(os.path.join(download_path, local_filename), 'wb') as file:
#                                         for chunk in response.iter_content(chunk_size=8192):
#                                             file.write(chunk)
#                                 except requests.exceptions.RequestException as e:
#                                     print(f"Failed to download {download_url}. Error: {e}")
#                                 time.sleep(1)
#                         break  # Exit the loop if success

#                     except StaleElementReferenceException:
#                         attempt += 1
#                         print(f"Attempt {attempt}/{max_attempts}: StaleElementReferenceException, retrying... {download_url}")

#                         if attempt == max_attempts:
#                             print("Max attempts reached, moving to the next category.")

#                 driver.get(base_url)
#                 WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'navigationDrawer')))

#             except (TimeoutException, ElementNotInteractableException, requests.HTTPError) as e:
#                 print(f"Error occurred: {category_url}")
#                 continue

def download_files(driver, download_path):
    base_urls = [
        'https://docs.oracle.com/en/cloud/saas/project-management/26a/oefpp/index.html',
        'https://docs.oracle.com/en/cloud/saas/financials/26a/oefbf/index.html',
        'https://docs.oracle.com/en/cloud/saas/procurement/26a/oefbp/index.html',
        'https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/26a/oefsc/index.html'
    ]

    session = requests.Session()

    for base_url in base_urls:
        driver.get(base_url)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'navigationDrawer')))

        # Expand all sections and get the link items
        section_items = driver.find_elements(By.CSS_SELECTOR, '#navigationDrawer li')

        for section in section_items:
            try:
                # Expand collapsible sections if there's an expand icon
                try:
                    expand_icon = section.find_element(By.CSS_SELECTOR, '.oj-clickable-icon-nocontext')
                    driver.execute_script("arguments[0].click();", expand_icon)
                    time.sleep(1)
                except:
                    pass  # It's not expandable — continue

                # Click the link inside the section
                links = section.find_elements(By.CSS_SELECTOR, 'a')
                for link in links:
                    link_text = link.text
                    try:
                        driver.execute_script("arguments[0].click();", link)
                        time.sleep(1)

                        # Wait for and collect any .xlsm download links
                        try:
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href$='.xlsm']"))
                            )
                            download_links = driver.find_elements(By.CSS_SELECTOR, "a[href$='.xlsm']")

                            for download_link in download_links:
                                download_url = download_link.get_attribute("href")
                                if download_url and download_url.endswith('.xlsm'):
                                    local_filename = download_url.split('/')[-1]
                                    print(f"Downloading: {local_filename}")
                                    try:
                                        # Copy cookies from Selenium to requests
                                        for cookie in driver.get_cookies():
                                            session.cookies.set(cookie['name'], cookie['value'])

                                        response = download_with_retry(session, download_url)
                                        with open(os.path.join(download_path, local_filename), 'wb') as file:
                                            for chunk in response.iter_content(chunk_size=8192):
                                                file.write(chunk)
                                        time.sleep(1)

                                    except requests.exceptions.RequestException as e:
                                        print(f"Failed to download {download_url}. Error: {e}")
                        except TimeoutException:
                            print(f"No downloads found in section: {link_text}")
                    except Exception as e:
                        print(f"Failed to click or process link: {link_text}. Error: {e}")
            except Exception as e:
                print(f"Error in section processing: {e}")

        print(f"Completed: {base_url}")


### Run Excel Macros on each file in folder
def run_excel_macros():

    with xw.App(visible=False) as app:
        wb = app.books.open('Clear_FBDIs - 20210412.xlsm')
        # Replace 'YourMacroName' with the name of your macro
        wb.macro('Sheet1.CommandButton1_Click')()
        wb.save()
        wb.close()

    with xw.App(visible=False) as app:
        wb = app.books.open('fbdi_compare.xlsm')
        # Replace 'YourMacroName' with the name of your macro
        app.macro('compare_fbdi_wrapper')()
        wb.save()
        wb.close()


### Execute

create_folders_and_clear()


# # Setup
chrome_options = Options()
chrome_options.add_argument('--ignore-ssl-errors=yes')
chrome_options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
download_path = './Originals'
if not os.path.exists(download_path):
    os.makedirs(download_path)

# Run the function
download_files(driver, download_path)

# Clean up
driver.quit()


# run_excel_macros()