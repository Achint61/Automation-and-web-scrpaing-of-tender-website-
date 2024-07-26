import os
import time
import random
import urllib.request
import pytesseract
from PIL import Image
import numpy as np
import cv2
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import csv
import string
import pandas

# Updated function to solve CAPTCHA using pytesseract with improved config for alphanumeric characters
def solve_captcha(image_path):
    captcha_image = Image.open(image_path)
    whitelist = string.ascii_uppercase + string.digits
    captcha_text = pytesseract.image_to_string(captcha_image, lang='eng', config='--psm 8 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')
    captcha_text = ''.join(filter(lambda char: char in whitelist, captcha_text))
    return captcha_text.strip()


# Function to handle retry mechanism for incorrect CAPTCHA

def handle_captcha_retry(driver, initial_url, captcha_input_xpath, submit_button_xpath):

    retry_count = 3
    while retry_count > 0:
        try:
            # if retry_count < 3:
            #     driver.get(initial_url)
            #     time.sleep(1)
            WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, 'captchaImage')))
            img_element = driver.find_element(By.ID, 'captchaImage')
            img_data = img_element.get_attribute('src')
            with urllib.request.urlopen(img_data) as response:
                data = response.read()
                with open("cap.png", mode="wb") as f:
                    f.write(data)
            captcha_text = solve_captcha("cap.png")
            print(f"Solved CAPTCHA: {captcha_text}")
            captcha_input = driver.find_element(By.XPATH, captcha_input_xpath)
            captcha_input.clear()
            captcha_input.send_keys(captcha_text)
            time.sleep(1)
            driver.find_element(By.XPATH, submit_button_xpath).click()
            WebDriverWait(driver, 1).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert_text = alert.text
            if "Please enter Captcha." in alert_text:
                print("CAPTCHA entered was incorrect. Retrying...")
                retry_count -= 1
                alert.accept()
                driver.find_element(By.XPATH,'//*[@id="captcha"]').click()
            else:
                print("Successfully passed CAPTCHA verification!")
                return True
        except TimeoutException:
            print("Successfully passed CAPTCHA verification!")
            return True
    print("Failed to pass CAPTCHA verification after retries.")
    return False

def parse_second_layer(driver,output_csv):
    global column_name
    Downloader_Check = True
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    tables = soup.find_all('table',class_='tablebg')
    allowed_index = [0,3,4,5,6]
    list_header = []
    tender_info = []
    templist = []
    for i in allowed_index:
        table = tables[i]
        title = table.find_all('td',class_='td_caption')
        for j in title:
            try:
                list_header.append(j.get_text())
            except:
                continue
        info = table.find_all('td',class_='td_field')
        for k in info:
            templist.append(k.get_text())
    list_header.extend(['Tender Link','Tender Document Download Link','Tender Documents ZIP Download Link'])
    templist.extend([driver.current_url])
    try:
        templist.extend([driver.find_element(By.XPATH,'//*[@id="DirectLink_0"]').get_attribute('href'),driver.find_element(By.XPATH,'//*[@id="DirectLink_7"]').get_attribute('href')])
    except:
        templist.extend(['NA','NA'])
        Downloader_Check = False
    tender_info.append(templist)
    data_frame = pandas.DataFrame(data = tender_info, columns = list_header)
    if(column_name):
        data_frame.to_csv(output_csv,mode='a',index = False)
        column_name = False  
    else:    
        data_frame.to_csv(output_csv,mode='a',index = False,header=False)  
    print("2nd Layer parsed successfully!")
    if(Downloader_Check):
        path = r"C:\Competetive Coding\Tender Docs"
        dir_list = sorted(os.listdir(path))
        try:
            os.rename(f"{path}/{dir_list[-1]}",f"{path}/{data_frame['Tender ID'].values[0]}.zip")
            os.rename(f"{path}/{dir_list[-2]}",f"{path}/{data_frame['Tender ID'].values[0]}.pdf")
        except:
            pass
    time.sleep(1)




# Function to parse table data and save to CSV with unique filename based on timestamp

def parse_and_save_table(driver, output_csv,final_csv):
    global captcha_check
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table[@class="list_table" and @id="table"]')))
        visited_urls = set()
        while True:
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            table = soup.find('table', class_='list_table', id='table')
            if not table:
                print("Table not found")
                return False
            current_url = driver.current_url
            if current_url in visited_urls:
                print("Already visited this page. Skipping...")
                break
            visited_urls.add(current_url)
            with open(output_csv, 'a', newline='', encoding='utf-8') as csvfile:
                csvwriter = csv.writer(csvfile)
                rows = table.find_all('tr')
                for row in rows:
                    cols = row.find_all(['th', 'td'])
                    cols = [col.get_text(strip=True) for col in cols]
                    csvwriter.writerow(cols)
            print(f"Table data saved to {output_csv}")
            #Storing XPath for all buttons for accessing layer 2
            download_links = ['//*[@id="DirectLink_0"]','//*[@id="DirectLink_0_0"]','//*[@id="DirectLink_0_1"]','//*[@id="DirectLink_0_2"]','//*[@id="DirectLink_0_3"]','//*[@id="DirectLink_0_4"]','//*[@id="DirectLink_0_5"]','//*[@id="DirectLink_0_6"]','//*[@id="DirectLink_0_7"]','//*[@id="DirectLink_0_8"]','//*[@id="DirectLink_0_9"]','//*[@id="DirectLink_0_10"]','//*[@id="DirectLink_0_11"]','//*[@id="DirectLink_0_12"]','//*[@id="DirectLink_0_13"]','//*[@id="DirectLink_0_14"]','//*[@id="DirectLink_0_15"]','//*[@id="DirectLink_0_16"]','//*[@id="DirectLink_0_17"]','//*[@id="DirectLink_0_18"]']
            for i in download_links:
                try:
                    driver.find_element(By.XPATH,i).click()
                    time.sleep(0.5)
                    if captcha_check:
                        driver.find_element(By.XPATH,'//*[@id="DirectLink_8"]').click()
                        time.sleep(0.5)
                        handle_captcha_retry(driver,initial_url,captcha_input_xpath,'//*[@id="Submit"]')
                        captcha_check = False
                        time.sleep(0.5)
                    try:
                        driver.find_element(By.XPATH,'//*[@id="DirectLink_7"]').click()
                        time.sleep(10)
                    except:
                        pass
                    try:
                        driver.find_element(By.XPATH,'//*[@id="DirectLink_0"]').click()
                        time.sleep(10)
                    except:
                        pass
                    parse_second_layer(driver,final_csv)
                    driver.find_element(By.XPATH,'//*[@id="DirectLink_11"]').click()
                except TimeoutException:
                    print("All Tender Documents Downloaded!")
                    break
            try:
                next_button = driver.find_element(By.XPATH, '//*[@id="linkFwd"]')
                if "disabled" in next_button.get_attribute("class"):
                    print("No more pages to load.")
                    break
                next_button.click()
                time.sleep(5)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table[@class="list_table" and @id="table"]')))
            except NoSuchElementException:
                print("Next button not found or no more pages.")
                break
        return True
    except TimeoutException:
        print("Timed out waiting for table to load")
        return False

if __name__ == '__main__':
    try:
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : r"C:\Competetive Coding\Tender Docs"}
        chrome_options.add_experimental_option('prefs', prefs)
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options= chrome_options)
        driver.get('https://etender.up.nic.in/nicgep/app')
        initial_url = driver.current_url
        time.sleep(1)
        link_xpath = "/html/body/table/tbody/tr[1]/td/table/tbody/tr[4]/td/table/tbody/tr/td[2]/span/span[1]/span[1]/a"
        driver.find_element(By.XPATH, link_xpath).click()
        time.sleep(1)
        tender_type_dropdown_xpath = "/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/form/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[5]/td/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select"
        tender_type_dropdown = Select(driver.find_element(By.XPATH, tender_type_dropdown_xpath))
        tender_type_choice = "Open Tender"
        tender_type_dropdown.select_by_visible_text(tender_type_choice)
        time.sleep(1)
        tender_category_dropdown_xpath = "/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/form/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[5]/td/table/tbody/tr/td/table/tbody/tr[6]/td[4]/select"
        tender_category_dropdown = Select(driver.find_element(By.XPATH, tender_category_dropdown_xpath))
        options = [option.text for option in tender_category_dropdown.options if option.text]
        selected_category = random.choice(options)
        tender_category_dropdown.select_by_visible_text(selected_category)
        print(f"Tender Type: {tender_type_choice}")
        print(f"Tender Category: {selected_category}")
        time.sleep(1)
        captcha_input_xpath = '//*[@id="captchaText"]'
        submit_button_xpath = '//*[@id="submit"]'
        captcha_passed = handle_captcha_retry(driver, initial_url, captcha_input_xpath, submit_button_xpath)
        if captcha_passed:

            print("CAPTCHA verification passed. Proceeding with further steps...")
            output_dir = "./output"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            output_csv = f"{output_dir}/table_data_{timestamp}.csv"
            final_csv = f"{output_dir}/tender_data_{timestamp}.csv"
            global captcha_check
            captcha_check = True
            global column_name
            column_name = True
            parse_and_save_table(driver, output_csv,final_csv)
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()

