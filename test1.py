import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, JavascriptException
from selenium.webdriver.chrome.options import Options

# Chrome options to speed up the browser
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--blink-settings=imagesEnabled=false")  # Disable images
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-css")  # Disable CSS rendering

# Initialize the Chrome WebDriver with optimized options
service = Service(r'C:\python\pythonProject\chromedriver-win64\chromedriver.exe')
driver = webdriver.Chrome(service=service, options=chrome_options)

# Set up WebDriver wait
wait = WebDriverWait(driver, 8)  # Reduced wait times for faster operation

# Open the website
driver.get("https://www.educacion.gob.es/centros/inicio#")
print("Website opened successfully")

# Click on "Todo el territorio"
todo_territorio = wait.until(EC.element_to_be_clickable((By.XPATH, "//li/a[@title='Todo el territorio']")))
todo_territorio.click()
print("Clicked on 'Todo el territorio'")

# Click on "Buscar" to retrieve results
buscar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='btn btn-primary btn-sm']")))
buscar_button.click()
print("Clicked on 'Buscar' button")

# Wait for the initial "Detalle centro" buttons to appear
wait.until(EC.presence_of_element_located((By.XPATH, "//td/a/i[@title='Detalle centro']")))
print("Initial 'Detalle centro' buttons found")

# CSV file path
csv_file_path = "result_fast.csv"

# Columns for the CSV file
columns = ["Código de centro", "Ubicación", "Identificación"]

# Check if CSV file exists, if not create it with headers
if not os.path.isfile(csv_file_path):
    pd.DataFrame(columns=columns).to_csv(csv_file_path, index=False)

# Set for visited URLs
visited_urls = set()

# Process entries continuously
while True:
    try:
        print("Finding table rows with 'Detalle centro' buttons...")

        # Fetch "Detalle centro" elements once and store them for processing
        detalle_buttons = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//td/a/i[@title='Detalle centro']")))

        for index, detalle_button in enumerate(detalle_buttons):
            try:
                # Get the parent link dynamically for each detail button
                detalle_link = detalle_button.find_element(By.XPATH, "./ancestor::a")
            except StaleElementReferenceException:
                print(f"Stale element detected on row {index + 1}, skipping...")
                continue  # Skip this element and move on to the next one

            # Get the 'onclick' value for dynamic JavaScript execution
            onclick_script = detalle_link.get_attribute('onclick')

            # Skip if this entry has already been processed
            if onclick_script in visited_urls:
                continue

            # Add the 'onclick' script to visited URLs set
            visited_urls.add(onclick_script)

            # Scroll to the button and execute the 'onclick' script
            driver.execute_script("arguments[0].scrollIntoView(true);", detalle_link)

            # Execute the dynamic JavaScript to load the details page
            try:
                print(f"Executing 'onclick' script for row {index + 1}: {onclick_script}")
                driver.execute_script(onclick_script)
            except JavascriptException as e:
                print(f"JavaScript execution failed for row {index + 1}: {str(e)}")
                continue  # Skip this entry if execution fails

            # Wait for the details page to load
            wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='col-md-6']")))
            print(f"Details page loaded for row {index + 1}")

            # Extract the required details
            try:
                codigo_centro = driver.find_element(By.XPATH, "/html/body/main/div/div[5]/div[1]").text
            except NoSuchElementException:
                codigo_centro = "N/A"

            try:
                ubicacion = driver.find_element(By.XPATH, "/html/body/main/div/div[7]").text
            except NoSuchElementException:
                ubicacion = "N/A"

            try:
                identificacion = driver.find_element(By.XPATH, "/html/body/main/div/div[5]").text
            except NoSuchElementException:
                identificacion = "N/A"

            # Create a dictionary with the extracted data
            data_dict = {
                "Código de centro": codigo_centro,
                "Ubicación": ubicacion,
                "Identificación": identificacion
            }

            # Create a DataFrame from the extracted data
            df = pd.DataFrame([data_dict])

            # Append the data to the CSV file
            df.to_csv(csv_file_path, index=False, mode='a', header=False)
            print(f"Data for row {index + 1} saved to CSV")

            # Navigate back using the "Atras" button
            print(f"Navigating back to the results page for row {index + 1}...")
            back_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div/a/i[@class='fas fa-arrow-circle-left']")))
            back_button.click()
            print(f"Navigated back to the results page for row {index + 1}")

            # Wait for the table to reload
            wait.until(EC.presence_of_element_located((By.XPATH, "//td/a/i[@title='Detalle centro']")))
            print(f"Results page reloaded for row {index + 1}")

    except (NoSuchElementException, TimeoutException, StaleElementReferenceException) as e:
        print(f"An error occurred: {e}")
        # Refresh the page to recover from the error
        driver.get("https://www.educacion.gob.es/centros/buscarCentros")
        wait.until(EC.presence_of_element_located((By.XPATH, "//td/a/i[@title='Detalle centro']")))
        print("Refreshed the page after error, continuing the process...")
        continue  # Continue with the next iteration after recovering from error

# Close the driver when done (this part will only run if you break out of the loop)
driver.quit()
