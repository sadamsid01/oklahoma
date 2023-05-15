import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import csv
import os

url = "https://verifyroofing.cib.ok.gov/roofing/search"
folder_location = os.getcwd()
chrome_driver = "chromedriver.exe"
DRIVER_PATH = os.path.join(folder_location, "chromedriver.exe")
print(DRIVER_PATH)
opts = ChromeOptions()
driver = webdriver.Chrome(executable_path=DRIVER_PATH, options=opts)
driver.maximize_window()
delay = 5
driver.get(url)

# Get Directory Location
folder_location = os.getcwd()
leads_folder = os.path.join(folder_location, "Leads")

# CSV header
header_list = ['Name', 'Registration', 'Type', 'Status', 'Commercial', 'Company', 'DBA', 'Address', 'Phone',
               'Expiration']
csv_name = '.csv'

# Save Business Info in CSV
csv_file_name = leads_folder + '/' + 'leads' + csv_name
f = open(csv_file_name, 'w',
         newline='')
writer = csv.writer(f)
writer.writerow(header_list)


def oklahoma():
    time.sleep(5)
    # Select Roofing Contract Registration
    try:
        select_option = WebDriverWait(driver, delay).until(
            EC.visibility_of_element_located((By.XPATH,
                                              '/html/body/app-root/app-main/div/app-content/app-search/div[1]/form/div[8]/select/option[3]')))
        if select_option:
            select_option.click()

        # Click on Submit Button
        try:
            submit_button = WebDriverWait(driver, delay).until(
                EC.visibility_of_element_located(
                    (By.XPATH, '/html/body/app-root/app-main/div/app-content/app-search/div[1]/form/div[9]/button')))
            if submit_button:
                submit_button.click()
                time.sleep(1)
                driver.execute_script("window.scrollBy(0, 700);")

                # get last result
                try:
                    last_results = WebDriverWait(driver, 100).until(
                        EC.element_to_be_clickable((By.XPATH,
                                                    '/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-footer/div/datatable-pager/ul/li[9]/a')))
                    if last_results:
                        time.sleep(1)
                        last_results.click()
                        # Get total results
                        try:
                            results = WebDriverWait(driver, delay).until(
                                EC.visibility_of_element_located((By.XPATH,
                                                                  '/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-footer/div/datatable-pager/ul/li[7]/a')))
                            if results:
                                results = results.text
                                results = int(results)
                                print(results)

                                # Go back to 1st order
                                try:
                                    first_result = WebDriverWait(driver, delay).until(
                                        EC.visibility_of_element_located((By.XPATH,
                                                                          '/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-footer/div/datatable-pager/ul/li[1]/a/i')))
                                    if first_result:
                                        # driver.execute_script("window.scrollBy(0, -300);")
                                        first_result.click()

                                    # Loop through Results
                                    for result in range(results):
                                        driver.execute_script("window.scrollBy(0, -200);")
                                        result += 1

                                        for row in range(7):
                                            data = []
                                            company_data = []
                                            row += 1

                                            # Check If business is In Good Standing or not
                                            try:
                                                check_status = WebDriverWait(driver, delay).until(
                                                    EC.visibility_of_element_located((By.XPATH,
                                                                                      f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/datatable-body-row/div[2]/datatable-body-cell[5]/div/p')))
                                                if check_status:
                                                    check_status = check_status.text
                                                    if check_status == 'In Good Standing':
                                                        print('this one')
                                                        # Click on business info
                                                        try:
                                                            click_business = WebDriverWait(driver, delay).until(
                                                                EC.visibility_of_element_located((By.XPATH,
                                                                                                  f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/datatable-body-row/div[2]/datatable-body-cell[1]/div')))
                                                            if click_business:
                                                                # driver.execute_script("arguments[0].scrollIntoView();", click_business)
                                                                click_business.click()
                                                                # Scrape business Name
                                                                try:
                                                                    name = WebDriverWait(driver, delay).until(
                                                                        EC.visibility_of_element_located((By.XPATH,
                                                                                                          f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/div/div/div[2]')))
                                                                    if name:
                                                                        name = name.text
                                                                        data.append(name)
                                                                except Exception as e:
                                                                    name = ''
                                                                    data.append(name)
                                                                    print('Name not found.')

                                                                # Count number of div to check if DBA is available or not
                                                                try:
                                                                    total_divs = WebDriverWait(driver, delay).until(
                                                                        EC.visibility_of_element_located((By.XPATH,
                                                                                                          f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/div/div')))
                                                                    total_divs = total_divs.text
                                                                    print(total_divs)
                                                                    total_divs = total_divs.split('\n')
                                                                    div_length = len(total_divs)
                                                                    name = total_divs[0]
                                                                    name = name.split(',')
                                                                    length = len(name)
                                                                    if length == 3:
                                                                        type = name[0]
                                                                        registration = name[1]
                                                                        registration = registration.split(':')
                                                                        registration = registration[1]
                                                                        status = 'In Good Standing'
                                                                        commercial = 'None'
                                                                        data.append(registration)
                                                                        data.append(type)
                                                                        data.append(status)
                                                                        data.append(commercial)
                                                                    else:
                                                                        type = name[0]
                                                                        registration = name[2]
                                                                        registration = registration.split(':')
                                                                        registration = registration[1]
                                                                        status = 'In Good Standing'
                                                                        commercial = 'None'
                                                                        data.append(registration)
                                                                        data.append(type)
                                                                        data.append(status)
                                                                        data.append(commercial)
                                                                    if div_length == 6:
                                                                        company = total_divs[2]
                                                                        dba = 'n/a'
                                                                        expiration_time = total_divs[3]
                                                                        expiration_time = expiration_time.split(':')
                                                                        expiration_time = expiration_time[1]
                                                                        address = total_divs[4]
                                                                        phone = total_divs[5]
                                                                        data.append(company)
                                                                        data.append(dba)
                                                                        data.append(address)
                                                                        data.append(phone)
                                                                        data.append(expiration_time)
                                                                        company_data.append(data)
                                                                        writer.writerows(company_data)

                                                                    if (div_length == 7):
                                                                        company = total_divs[2]
                                                                        dba = total_divs[3]
                                                                        expiration_time = total_divs[4]
                                                                        expiration_time = expiration_time.split(':')
                                                                        expiration_time = expiration_time[1]
                                                                        address = total_divs[5]
                                                                        phone = total_divs[6]
                                                                        data.append(company)
                                                                        data.append(dba)
                                                                        data.append(address)
                                                                        data.append(phone)
                                                                        data.append(expiration_time)
                                                                        company_data.append(data)
                                                                        writer.writerows(company_data)
                                                                except Exception as e:
                                                                    print(e)
                                                                    print('error in counting divs')

                                                                # try:
                                                                #     name = WebDriverWait(driver, delay).until(
                                                                #         EC.visibility_of_element_located(
                                                                #             (By.XPATH,
                                                                #              f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/div/div/div[1]/strong')))
                                                                #     if name:
                                                                #         name = name.text
                                                                #         name = name.split(',')
                                                                #         length = len(name)
                                                                #         if length == 3:
                                                                #             type = name[0]
                                                                #             registration = name[1]
                                                                #             registration = registration.split(':')
                                                                #             registration = registration[1]
                                                                #             status = 'In Good Standing'
                                                                #             commercial = 'None'
                                                                #             data.append(registration)
                                                                #             data.append(type)
                                                                #             data.append(status)
                                                                #             data.append(commercial)
                                                                #         else:
                                                                #             type = name[0]
                                                                #             registration = name[2]
                                                                #             registration = registration.split(':')
                                                                #             registration = registration[1]
                                                                #             status = 'In Good Standing'
                                                                #             commercial = 'None'
                                                                #             data.append(registration)
                                                                #             data.append(type)
                                                                #             data.append(status)
                                                                #             data.append(commercial)
                                                                #
                                                                # except Exception as e:
                                                                #     type = ''
                                                                #     registration = ''
                                                                #     status = 'In Good Standing'
                                                                #     commercial = 'None'
                                                                #     data.append(registration)
                                                                #     data.append(type)
                                                                #     data.append(status)
                                                                #     data.append(commercial)
                                                                #     print('Error in Getting Name and Registration.')
                                                                # # Scrape Company Name
                                                                # # try:
                                                                # #     company = WebDriverWait(driver, delay).until(
                                                                # #         EC.visibility_of_element_located(
                                                                # #             (By.XPATH,
                                                                # #              f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{result}]/div/div/div[3]')))
                                                                # #     if company:
                                                                # #         company = company.text
                                                                # #         data.append(company)
                                                                # # except Exception as e:
                                                                # #     company = ''
                                                                # #     data.append(company)
                                                                # #     print('Error in Getting company.')
                                                                #
                                                                # # Scrape Company Name
                                                                # try:
                                                                #     company = WebDriverWait(driver, delay).until(
                                                                #         EC.visibility_of_element_located(
                                                                #             (By.XPATH,
                                                                #              f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{result}]/div/div/div[3]')))
                                                                #     if company:
                                                                #         company = company.text
                                                                #         dba = 'n/a'
                                                                #         data.append(company)
                                                                #         data.append(dba)
                                                                #
                                                                # except Exception as e:
                                                                #     company = ''
                                                                #     dba = 'n/a'
                                                                #     data.append(company)
                                                                #     data.append(dba)
                                                                #     print('Error in Getting company.')
                                                                #
                                                                # # Scrape business Address
                                                                # try:
                                                                #     address = WebDriverWait(driver, delay).until(
                                                                #         EC.visibility_of_element_located(
                                                                #             (By.XPATH,
                                                                #              f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/div/div/div[6]')))
                                                                #     if address:
                                                                #         address = address.text
                                                                #         data.append(address)
                                                                # except Exception as e:
                                                                #     address = ''
                                                                #     data.append(address)
                                                                #     print('Address not found')
                                                                #
                                                                # # Scrape business phone
                                                                # try:
                                                                #     phone = WebDriverWait(driver, delay).until(
                                                                #         EC.visibility_of_element_located(
                                                                #             (By.XPATH,
                                                                #              f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/div/div/div[7]')))
                                                                #     if phone:
                                                                #         phone = phone.text
                                                                #         data.append(phone)
                                                                # except Exception as e:
                                                                #     phone = ''
                                                                #     data.append(phone)
                                                                #     print('Phone Number  not found')
                                                                #
                                                                # # Scrape business Expiration Time
                                                                # try:
                                                                #     expiration_time = WebDriverWait(driver,
                                                                #                                     delay).until(
                                                                #         EC.visibility_of_element_located(
                                                                #             (By.XPATH,
                                                                #              f'/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper[{row}]/div/div/div[5]')))
                                                                #     if expiration_time:
                                                                #         expiration_time = expiration_time.text
                                                                #         expiration_time = expiration_time.split(':')
                                                                #         expiration_time = expiration_time[1]
                                                                #         data.append(expiration_time)
                                                                #         company_data.append(data)
                                                                #         writer.writerows(company_data)
                                                                #         # time.sleep(0.5)
                                                                #         # driver.back()
                                                                # except Exception as e:
                                                                #     expiration_time = ''
                                                                #     data.append(expiration_time)
                                                                #     company_data.append(data)
                                                                #     writer.writerows(company_data)
                                                                #     print('Expiration Time  not found')
                                                            click_business.click()
                                                        except Exception as e:
                                                            print('Error on Clicking Business Information')
                                                    else:
                                                        print('Business is not In Good Running')
                                                        driver.execute_script("window.scrollBy(0, 50);")
                                            except Exception as e:
                                                print('Error in Getting status information')

                                        # Go to next tab results
                                        try:
                                            next_tab = WebDriverWait(driver, delay).until(
                                                EC.visibility_of_element_located((By.XPATH,
                                                                                  '/html/body/app-root/app-main/div/app-content/app-search/div[2]/ngx-datatable/div/datatable-footer/div/datatable-pager/ul/li[8]/a/i')))
                                            if next_tab:
                                                driver.execute_script("window.scrollBy(0, 500);")
                                                next_tab.click()
                                        except Exception as e:
                                            print('Error in going to next tab results')
                                except Exception as e:
                                    print('Error in going back to 1st result')
                        except Exception as e:
                            print('Error in getting total tesults')
                except Exception as e:
                    print('Error in getting last result')
        except Exception as e:
            print('Error on Clicking Submitting button')

    except Exception as e:
        print('Error on Selecting Roofing Contract Registration')


def remove_duplicates(csv_file_name):
    # Read the CSV file into a list of lists
    with open(csv_file_name, 'r', newline='') as f:
        reader = csv.reader(f)
        # Get the header row
        header = next(reader)
        data = [row for row in reader]
    # Remove duplicates from the list of rows
    unique_data = [header]
    for row in data:
        if row not in unique_data:
            unique_data.append(row)
    # Overwrite the original CSV file with the updated data
    with open(csv_file_name, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(unique_data)


oklahoma()
f.close()
# remove_duplicates(csv_file_name)
driver.quit()
