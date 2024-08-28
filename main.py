from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium import webdriver
import os
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import Select
import pandas as pd


Options = webdriver.ChromeOptions()
Options.add_argument('--no-sandbox')
Options.add_argument('--disable-dev-shm-usage')
Options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36')
Options.add_argument('--start-maximized')
Options.add_argument('--headless=new')

service = ChromeService(ChromeDriverManager().install())

driver = webdriver.Chrome(service=service, options=Options)


def appendProduct(file_path2, data):
    temp_file = 'temp_file.xlsx'
    if os.path.isfile(file_path2):
        df = pd.read_excel(file_path2, engine='openpyxl')
    else:
        df = pd.DataFrame()
    df_new_row = pd.DataFrame([data])
    df = pd.concat([df, df_new_row], ignore_index=True) 
    try:
        df.to_excel(temp_file, index=False, engine='openpyxl')
    except Exception as e:
        print(f"An error occurred while saving the temporary file: {str(e)}")
        return False 
    try:
        os.replace(temp_file, file_path2)
    except Exception as e:
        print(f"An error occurred while replacing the original file: {str(e)}")
        return False
    
    return True



def save_current_page(page_index):
    with open("current_page.txt", "w") as file:
        file.write(str(page_index))

def load_current_page():
    if os.path.exists("current_page.txt"):
        with open("current_page.txt", "r") as file:
            return int(file.read())
    return 1  


def click_next_page(times):
    for _ in range(times):
        try:
            next_button = driver.find_element(By.XPATH, "//img[@alt='next']")
            driver.execute_script("arguments[0].scrollIntoView();", next_button)
            driver.execute_script("window.scrollBy(0, -150)")
            driver.execute_script("arguments[0].click();", next_button)
            print(f"Page {_ + 1} clicked")
            try:
                searching = driver.find_element(By.XPATH, "//div[@class='searching']")
                if searching.is_displayed():
                    time.sleep(15)  
            except:
                time.sleep(1)
                pass 
        except Exception as e:
            print(f"Error clicking next page: {e}")



def get_data():
    driver.get("https://gateway.insurance.ohio.gov/UI/ODI.Agent.Public.UI/AgentSearch.mvc/DisplaySearch")
    time.sleep(2)
    
    license_type = driver.find_element(By.XPATH, "//select[@id='LicenseType']")
    license_selector = Select(license_type)
    license_selector.select_by_index(3)
    time.sleep(2)
    
    search_button = driver.find_element(By.XPATH, "//input[@value='Search']")
    search_button.click()
    time.sleep(20)
    
    current_page_index = load_current_page()
    last_page_index = 19694

    if current_page_index > 1:
        click_next_page(current_page_index - 1)

    while current_page_index <= last_page_index:
        agent_names = driver.find_elements(By.XPATH, "//li[contains(@class, 'agentName')]")
        address_1 = driver.find_elements(By.XPATH, "//ul[@class='agentDetails']//li[2]")
        address_2 = driver.find_elements(By.XPATH, "//ul[@class='agentDetails']//li[last()]")
        phone_numbers = driver.find_elements(By.XPATH, "//tr//span[@data-bind='text: odi.utils.formatPhoneNumber(agentPhone)']")

        for agent_name in agent_names:
            full_name = agent_name.text
            if ',' in full_name:
                last_name, first_name = [name.strip() for name in full_name.split(',', 1)]
            else:
                first_name= full_name.strip()
                last_name = ""
            parent_tr = agent_name.find_element(By.XPATH, "./ancestor::tr[1]")
            view_profile = parent_tr.find_element(By.XPATH, ".//a[@title='View Agent Profile']")
            view_profile.click()
            try:
                driver.switch_to.window(driver.window_handles[1])
                time.sleep(2)
                mail_button = driver.find_element(By.XPATH, "//a[.='Business Email Address']")
                mail_button.click()
                time.sleep(1)
                email = driver.find_element(By.XPATH, "//a[text()='Business Email Address']/parent::div/following-sibling::div[1]").text
            except:
                email = ""
            try:
                load_appointments=driver.find_element(By.XPATH, "//input[@value='Click to Load']")
                driver.execute_script("arguments[0].click();", load_appointments)
                first_td_element=WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//td[normalize-space(text())='Life']")))
                if first_td_element:
                    processed_companies=set()
                    life_td_elements = driver.find_elements(By.XPATH, "//td[normalize-space(text())='Life']")
                    for life_td in life_td_elements:
                        company_name = life_td.find_element(By.XPATH, "preceding-sibling::td[1]").text
                        if company_name not in processed_companies:
                            processed_companies.add(company_name)
                else:
                    processed_companies = set()
                    processed_companies.add(" ")
            except:
                processed_companies = set()
                processed_companies.add(" ")
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            try:
                address = address_1[agent_names.index(agent_name)].text + " " + address_2[agent_names.index(agent_name)].text
            except:
                address = ""
            try:
                phone_number = phone_numbers[agent_names.index(agent_name)].text
            except:
                phone_number = ""
            data = {
                "First Name": first_name,
                "Last Name": last_name,
                "Address": address,
                "Phone Number": phone_number,
                "Email": email.strip(),
                "Active Appointing Companies - Life": ", ".join(processed_companies)
            }
            print(data)
            if os.path.exists("insurances.xlsx"):
                df = pd.read_excel("insurances.xlsx")
                df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
            else:
                df = pd.DataFrame([data])
            
            try:
                df.to_excel("insurances.xlsx", index=False)
            except Exception as e:
                print(f"An error occurred while saving the file: {str(e)}")
        
        try:
            next_button = driver.find_element(By.XPATH, "//img[@alt='next']")
            next_button.click()
            try:
                searching = driver.find_element(By.XPATH, "//div[@class='searching']")
                if searching.is_displayed():
                    time.sleep(15)  
            except:
                time.sleep(2)
                pass  
            
           
            save_current_page(current_page_index)  
            current_page_index += 1
            
        except Exception as e:
            print(f"Error on page {current_page_index}: {e}")
            break


        
        

get_data()

