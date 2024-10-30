import os
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, StaleElementReferenceException,
    UnexpectedAlertPresentException, NoAlertPresentException
)
import time
import random
import re  # Import regex module to handle phone number formatting


# Function to set up Selenium WebDriver with random user agent
def setup_selenium():
    options = uc.ChromeOptions()

    # Set a static user-agent
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'
    ]
    user_agent = random.choice(user_agents)
    options.add_argument(f"user-agent={user_agent}")

    # Use undetected chromedriver to bypass detection
    driver = uc.Chrome(options=options)
    return driver


# Function to introduce random delays
def random_delay(min_delay=2, max_delay=5):
    time.sleep(random.uniform(min_delay, max_delay))


# Function to handle alerts (CAPTCHAs or popups)
def handle_alert(driver):
    try:
        alert = driver.switch_to.alert
        print(f"Unexpected alert detected: {alert.text}")
        alert.accept()
        print("Alert dismissed")
    except NoAlertPresentException:
        print("No alert present to handle.")


# Function to scroll randomly to simulate human activity
def simulate_human_activity(driver):
    # Scroll down randomly
    for _ in range(random.randint(1, 3)):
        scroll_by = random.randint(200, 600)
        driver.execute_script(f"window.scrollBy(0, {scroll_by});")
        random_delay(1, 3)


# Function to format and clean phone numbers
def format_phone_number(phone_number):
    # Extract numbers and ignore other symbols
    numbers = re.findall(r'[0-9]+', phone_number)
    cleaned_number = " ".join(numbers)  # Join them as space-separated digits
    return cleaned_number


# Function to save data to CSV properly without overwriting
def save_to_csv(data, csv_file_path):
    # Check if file exists to avoid overwriting headers
    file_exists = os.path.isfile(csv_file_path)

    # Write to the CSV file
    df = pd.DataFrame([data])
    df.to_csv(csv_file_path, index=False, mode='a', header=not file_exists)


# Initialize WebDriver
driver = setup_selenium()
wait = WebDriverWait(driver, 30)

# Open the webpage
driver.get("https://www.maharashtradirectory.com/product/rubber-sheet.html")
print("Website opened successfully")

# Scroll down the page to simulate human activity
simulate_human_activity(driver)

# CSV file path
csv_file_path = "rubber_sheet_vendor1.csv"
columns = ["Company Name", "Contact Person", "Telephone", "Address", "Product Details", "Activities"]

# Ensure we create the CSV file with headers only once
if not os.path.isfile(csv_file_path):
    pd.DataFrame(columns=columns).to_csv(csv_file_path, index=False)

# List to store visited URLs
visited_urls = set()

# Retry logic for failed attempts
max_attempts = 3

while True:
    try:
        print("Fetching company links...")

        # Fetch the list of companies by their clickable href elements (links)
        company_links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div/a[contains(@href, 'companyinfo')]")))

        # Iterate over all companies dynamically by re-fetching each time to avoid stale elements
        total_companies = len(company_links)

        for index in range(total_companies):
            attempt = 0
            while attempt < max_attempts:
                try:
                    # Re-fetch the company links dynamically for each iteration
                    company_links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div/a[contains(@href, 'companyinfo')]")))
                    company_link = company_links[index]

                    company_href = company_link.get_attribute('href')

                    # Skip if this entry has already been visited
                    if company_href in visited_urls:
                        break

                    # Add the href to visited URLs
                    visited_urls.add(company_href)

                    # Use JavaScript to click the company link
                    print(f"Clicking on company link : {company_href}")
                    driver.execute_script("arguments[0].click();", company_link)
                    random_delay(3, 5)

                    # Handle unexpected alerts such as CAPTCHA
                    handle_alert(driver)

                    # Wait for the details page to load
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "col-md-12")))
                    print(f"Details page loaded for company")

                    # Extract the individual fields separately using indexed XPath expressions

                    # Step 1: Extract Company Name (first block)
                    try:
                        company_name = driver.find_element(By.XPATH, "(//div[@class='view-details-search-matter'])[1]").text
                    except NoSuchElementException:
                        company_name = "N/A"

                    # Step 2: Extract Contact Person (second block with assumption)
                    try:
                        contact_person = driver.find_element(By.XPATH, "(//div[@class='view-details-search-matter'])[2]").text
                    except NoSuchElementException:
                        contact_person = "N/A"

                    # Step 3: Extract Telephone (third block where 'Tel No.' is included)
                    try:
                        telephone = driver.find_element(By.XPATH, "(//div[@class='view-details-search-matter'])[3]").text
                        telephone = telephone.replace("Tel No.", "").strip()
                        telephone = format_phone_number(telephone)  # Clean the phone number properly
                    except NoSuchElementException:
                        telephone = "N/A"

                    # Step 4: Extract Address (unique XPath)
                    try:
                        address = driver.find_element(By.XPATH, "/html/body/div[7]/div/div[1]/div/div/div/div[3]/div/div/div[2]/div/div/div[2]").text
                    except NoSuchElementException:
                        address = "N/A"

                    # Step 5: Extract Product Details (unique XPath)
                    try:
                        product_details = driver.find_element(By.XPATH, "/html/body/div[7]/div/div[1]/div/div/div/div[3]/div/div/div[5]").text
                    except NoSuchElementException:
                        product_details = "N/A"

                    # Step 6: Extract Activities (using provided XPath)
                    try:
                        activities = driver.find_element(By.XPATH, "/html/body/div[7]/div/div[1]/div/div/div/div[3]/div/div/div[4]/div[2]/div").text
                    except NoSuchElementException:
                        activities = "N/A"

                    # Save the extracted data to CSV
                    data = {
                        "Company Name": company_name,
                        "Contact Person": contact_person,
                        "Telephone": telephone,
                        "Address": address,
                        "Product Details": product_details,
                        "Activities": activities  # Adding activities
                    }

                    save_to_csv(data, csv_file_path)

                    print("Data for company  saved to CSV")

                    # Navigate back to the results page
                    print("Navigating back to the results page for company")
                    driver.back()

                    # Wait until the page is reloaded and continue
                    wait.until(EC.presence_of_element_located((By.XPATH, "//div/a[contains(@href, 'companyinfo')]")))
                    print(f"Results page reloaded for company")

                    break  # Exit retry loop after success

                except StaleElementReferenceException as e:
                    print("Stale element reference error for company:")
                    attempt += 1
                    continue  # Retry the loop for this company link

    except UnexpectedAlertPresentException as e:
        print("Unexpected alert detected: Dismissing the alert...")
        handle_alert(driver)
        driver.refresh()  # Refresh after handling the alert
        continue  # Retry the loop after handling

    except (NoSuchElementException, TimeoutException) as e:
        print(f"An error occurred: {e}")
        driver.refresh()  # Refresh the page to recover from the error
        wait.until(EC.presence_of_element_located((By.XPATH, "//div/a[@href]")))
        continue
        driver.quit()
