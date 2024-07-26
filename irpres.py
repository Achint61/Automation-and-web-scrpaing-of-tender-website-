import os
import time
import pytesseract
import csv
import random
import string
from PIL import Image, ImageEnhance, ImageFilter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pymupdf
import pandas

# Set the path to Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to preprocess and solve CAPTCHA using pytesseract
def solve_captcha(image_path):
    captcha_image = Image.open(image_path)

    # Convert the image to grayscale
    captcha_image = captcha_image.convert('L')

    # Enhance image contrast
    enhancer = ImageEnhance.Contrast(captcha_image)
    captcha_image = enhancer.enhance(2.0)  # Increase contrast

    # Apply image filters
    captcha_image = captcha_image.filter(ImageFilter.SMOOTH)
    captcha_image = captcha_image.filter(ImageFilter.SHARPEN)

    # Apply threshold to binarize the image
    threshold = 150
    captcha_image = captcha_image.point(lambda p: p > threshold and 255)

    # Additional preprocessing steps to remove noise and improve OCR accuracy
    captcha_image = captcha_image.filter(ImageFilter.MedianFilter())
    captcha_image = captcha_image.filter(ImageFilter.EDGE_ENHANCE)
    captcha_image = captcha_image.filter(ImageFilter.MaxFilter(3))

    # Use pytesseract to OCR the image, specifying characters
    captcha_text = pytesseract.image_to_string(captcha_image, lang='eng', config='--psm 8 --oem 3 -c tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyz')

    # Remove non-alphanumeric characters
    captcha_text = ''.join(filter(lambda char: char.isalnum(), captcha_text))

    return captcha_text.strip()

# Function to download table data from the current page
def download_table_data(driver, table_tag, writer):
    try:
        # Wait for the table to be present and visible
        table = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.TAG_NAME, table_tag))
        )

        # Find all rows in the table
        rows = table.find_elements(By.TAG_NAME, 'tr')

        # Process each row
        for row in rows:
            try:
                # Extract tender details from each row
                data_row = [data.text.strip() for data in row.find_elements(By.TAG_NAME, 'td')]
                if not data_row:
                    continue

                # Find the second layer link for each row
                second_layer_link = row.find_element(By.XPATH, './td[8]/a')
                pdf_url = download_pdf_from_link(driver, second_layer_link)

                # Append the PDF URL to the tender details
                data_row.append(pdf_url)

                # Write the combined data to the CSV file
                writer.writerow(data_row)
                print(f"Data written to CSV: {data_row}")

            except NoSuchElementException:
                print("No second layer link found in this row.")
                continue

            except Exception as e:
                print(f"Error processing row: {e}")
                continue

    except TimeoutException:
        print("Timeout waiting for table to load.")
    except Exception as e:
        print(f"Error downloading table data: {e}")

# Function to handle PDF download from a specific link
def download_pdf_from_link(driver, link):
    global second_layer
    global isfirst
    try:
        # Click on the link to open the new page
        link.click()
        time.sleep(2)  # Adjust sleep time as needed

        # Switch to the new window/tab opened
        driver.switch_to.window(driver.window_handles[-1])

        # Click the link to download the PDF
        pdf_download_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'styled-button-8'))
        )
        pdf_download_link.click()
        print("PDF download link clicked.")

        time.sleep(2)  # Wait for download to complete (adjust as needed)

        # Get the URL of the downloaded PDF
        pdf_url = driver.current_url

        # Get Tender ID for filename
        tender_id = driver.find_element(By.XPATH,"/html/body/table/tbody/tr[4]/td[2]/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/table[1]/tbody/tr[1]/td[2]").text

        # Close the new window/tab
        driver.close()

        # Switch back to the original window/tab
        driver.switch_to.window(driver.window_handles[0])

        #Extracting Data from PDF
        path = r"C:\Competetive Coding\Tender Docs"
        dir_list = sorted(os.listdir(path))
        os.rename(f"{path}\\{dir_list[-1]}",f"{path}\\{tender_id}.pdf")
        file = f"{path}\\{tender_id}.pdf"
        count = 0
        pdf = pymupdf.open(file)
        page = pdf[0]
        table = page.find_tables()[0].extract()
        del table[-2]
        Col = []
        Datalist = []
        templist = []
        count = 0
        for row in table:
            if(count == 5 or count == 7 or count == 9 or count == 12):
                Col.append(row[0])
                content = row[1].replace('\n',' ')
                templist.append(content)
            else:
                Col.extend([row[0],row[2]])
                content1 = row[1].replace('\n',' ')
                content2 = row[3].replace('\n',' ')
                templist.extend([content1,content2])
            count += 1
        
        #Storing extracted PDF data to a CSV file
        Datalist.append(templist)
        data_frame = pandas.DataFrame(data= Datalist, columns= Col)
        if(isfirst):
            data_frame.to_csv(second_layer,mode='a',index=False)
            isfirst = False
        else:
            data_frame.to_csv(second_layer,mode='a',index=False,header=False)

        return pdf_url

    except Exception as e:
        print(f"Error downloading PDF: {e}")
        return ''

chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory': r"C:\Competetive Coding\Tender Docs", "download.prompt_for_download": False, "download.directory_upgrade": True, "plugins.always_open_pdf_externally": True}
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(options=chrome_options)

try:
    # Open the website
    driver.get('https://www.ireps.gov.in/')

    # Wait until the close button is clickable and then click it
    close_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div/table/tbody/tr[3]/td/button'))
    )
    close_button.click()
    print("Close button clicked.")

    # Wait until the "Search e Tenders" button is present in the DOM
    search_e_tenders_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/section/div/div/div[1]/div/div[2]/div/div/div[1]/div/div[3]/ul/li[1]/a'))
    )

    # Scroll to the "Search e Tenders" button
    driver.execute_script("arguments[0].scrollIntoView();", search_e_tenders_button)

    # Wait until the "Search e Tenders" button is clickable and then click it using JavaScript
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/section/div/div/div[1]/div/div[2]/div/div/div[1]/div/div[3]/ul/li[1]/a'))
    )
    driver.execute_script("arguments[0].click();", search_e_tenders_button)
    print("Search e Tenders button clicked.")

    # Wait until the phone number input field is present in the DOM
    phone_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[5]/div[2]/div/div/form/div/table/tbody/tr[2]/td[3]/input'))
    )

    # Enter a phone number into the input field
    phone_input.send_keys('8826177337')  # Use a test number or temporary number
    print("Phone number entered.")

    # Handle CAPTCHA
    captcha_img_xpath = '//*[@id="verimage"]/img[1]'
    captcha_input_xpath = '//*[@id="verification"]'

    # Wait for CAPTCHA image to load and solve it
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, captcha_img_xpath)))
    captcha_img = driver.find_element(By.XPATH, captcha_img_xpath)
    img_data = captcha_img.screenshot_as_png
    with open("cap.png", mode="wb") as f:
        f.write(img_data)

    captcha_text = solve_captcha("cap.png")
    print(f"Solved CAPTCHA: {captcha_text}")

    # Enter CAPTCHA text
    captcha_input = driver.find_element(By.XPATH, captcha_input_xpath)
    captcha_input.clear()
    captcha_input.send_keys(captcha_text)
    time.sleep(1)

    # Wait for OTP input field to be present (indicating CAPTCHA verification success)
    otp_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'otp'))
    )

    # Enter OTP into the OTP input field
    otp_input.send_keys("980893")  # Enter the OTP directly here
    print("Entered OTP.")

    # Click the Proceed button
    proceed_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[2]/div/div/form/div/input[1]'))
    )
    proceed_button.click()
    print("Proceed button clicked.")

    # Click on the "Active Tenders" button
    active_tenders_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="activeTenderId"]'))
    )
    active_tenders_button.click()
    print("Active Tenders button clicked.")

    # Define table tag for the initial table
    table_tag = 'table'

    # Generate a random filename for the CSV
    output_dir = "C:/Competetive Coding/output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    filename = f"{output_dir}/first_layer{timestamp}.csv"
    global second_layer
    second_layer = f"{output_dir}/second_layer{timestamp}.csv"
    global isfirst
    isfirst = True
    # Open the CSV file for writing data rows and PDF URLs
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # Download data from the initial page
        download_table_data(driver, table_tag, writer)
        print("Data from the initial page downloaded.")

        # Navigate through other pages
        page_number = 2  # Start from the second page
        while True:
            try:
                # Find the next page button element by using its XPath
                next_page_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f'//a[text()="{page_number}"]'))
                )

                # Download data from the current page
                download_table_data(driver, table_tag, writer)
                print(f"Data from page {page_number} downloaded.")

                # Click on the next page button
                next_page_button.click()
                print(f"Clicked on page {page_number}.")

                page_number += 1  # Move to the next page

            except NoSuchElementException:
                print(f"No more pages available. Finished scraping.")
                break

            except TimeoutException:
                print(f"Timeout waiting for page {page_number} to load.")
                break

            except Exception as e:
                print(f"Error navigating through page {page_number}: {e}")
                break

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the WebDriver
    driver.quit()
