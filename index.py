import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
import urllib

import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = Options()

# Initialize an instance of the Chrome driver (browser)
options.page_load_strategy = 'eager'
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)
flat_data = []

try:
    print("home page")
    driver.get("https://www.hoamanagement.com/hoa-resource-search")
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, "//h3[contains(text(), 'Search by City')]")))

    listTable = driver.find_elements(
        By.XPATH, "//ul[@class='citylist']/li/a")

    for id, cityData in enumerate(listTable):
        print(cityData.text)
        cityName = cityData.text
        link = cityData.get_attribute("href")
        print("Go to link ", link)

        # Store the current window handle
        main_window = driver.current_window_handle
        driver.execute_script("window.open('', '_blank');")  # Open a new tab
        # Switch to the new tab
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(link)
        try:
            listCompany = driver.find_elements(
                By.XPATH, "//div[@class='listing']/div[@class='row']/div")
            print(len(listTable), len(listCompany))
            for index, companyData in enumerate(listCompany):
                # print(companyData.text)
                print("Current data : ", id, ":", index)
                try:
                    hidden_span = companyData.find_element(By.TAG_NAME, "span")
                    companyName = hidden_span.get_attribute("textContent")

                    print("companyName ", companyName)
                    address = cityName.split(",")
                    state = companyName
                    try:
                        phone = companyData.find_element(
                            By.XPATH, "./div/div[@class='company-contact-number']/h4/a").text
                    except NoSuchElementException:
                        phone = ""
                    sub_link_div = companyData.find_element(
                        By.XPATH, "./div/a[@class='btn btn-blue city contact-management-btn']")
                    sub_link_href = sub_link_div.get_attribute("href")
                    infoData ={
                        "City": address[0],
                        "State": address[1],
                        "Company Email": "",
                        "Company Name": companyName,
                        "Company Number": phone,
                        "Team Member 1 Name": "",
                        "Team Member 1 Position": "",
                        "Team Member 2 Name": "",
                        "Team Member 2 Position": ""}
                    sub_main_window = driver.current_window_handle
                    driver.execute_script(
                        "window.open('', '_blank');")  # Open a new tab
                    # Switch to the new tab
                    driver.switch_to.window(driver.window_handles[-1])
                    print("sub_link_href",sub_link_href)
                    driver.get(sub_link_href)
                    # After you are done with the new tab, close it
                    try:
                        website_div = driver.find_element(
                            By.LINK_TEXT, "View Website")
                        website = website_div.get_attribute("href")
                    except NoSuchElementException:
                        website = ""
                    infoData["Website"] =website
                    driver.close()
                    # Switch back to the main tab
                    driver.switch_to.window(sub_main_window)
                    flat_data.append(infoData)
                except NoSuchElementException:
                    print("No company name")
        except NoSuchElementException:
            print("Invalid list company")

        # After you are done with the new tab, close it
        driver.close()
        # Switch back to the main tab
        driver.switch_to.window(main_window)
except NoSuchElementException:
    print("Not found home page")

df = pd.DataFrame(flat_data)
excel_writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
df.to_excel(excel_writer, sheet_name='Sheet1', index=False)

# Save the Excel file
excel_writer.close()
