import subprocess
import sys

# 필요한 패키지가 설치되어 있는지 확인
# try:
#     subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium", "webdriver_manager", "pandas", "openpyxl", "tqdm", "keyboard"])
# except subprocess.CalledProcessError as e:
#     print(f"Error installing packages: {e}")
#     sys.exit(1)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import openpyxl
from openpyxl.styles import Font
from tqdm import tqdm
import logging
import os
from selenium.webdriver.chrome.service import Service
import keyboard  # 키보드 입력을 감지하기 위한 라이브러리

# 로깅 설정
logging.basicConfig(filename='mwc_crawler.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def setup_driver():
    # Chrome WebDriver 설정
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def wait_for_element(driver, selector, by=By.CSS_SELECTOR, timeout=10):
    # 특정 요소가 나타날 때까지 대기
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
        return element
    except TimeoutException:
        logging.warning(f"Element not found: {selector}")
        return None

def wait_for_page_load(driver, timeout=30):  # 대기 시간을 30초로 늘림
    # 페이지가 완전히 로드될 때까지 대기
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.exhibitor-card a"))
        )
    except TimeoutException:
        logging.warning("Page did not load properly")

def get_text_or_null(driver, xpath):
    try:
        element = driver.find_element(By.XPATH, xpath)
        return element.text.strip() if element else None
    except NoSuchElementException:
        return None

def get_company_details(driver, company_url):
    try:
        driver.get(company_url)
        
        exhibitor = get_text_or_null(driver, '//*[@id="headerContainer"]/div/div[3]/nav/ul/li[5]')
        
        exhibitor_header = [
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[1]/span'),
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[2]/span'),
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[3]/span'),
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[4]/span')
        ]
        
        information = get_text_or_null(driver, '//*[@id="maincontent"]/div')
        
        link = [None] * 6
        location = [None] * 6
        interests = [None] * 6
        
        aside_elements = driver.find_elements(By.XPATH, '//*[@id="exhibitor-container"]/aside/div')
        for i, element in enumerate(aside_elements, start=1):
            heading = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{i}]/h5')
            if heading == "Contacts & Links":
                for j in range(1, 7):
                    link[j-1] = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{i}]/ul/li[{j}]/a')
            elif heading == "Location":
                for j in range(1, 7):
                    location[j-1] = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{i}]/ul/li[{j}]/a')
            elif heading == "Interests":
                for j in range(1, 7):
                    interests[j-1] = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{i}]/ul/li[{j}]')
        
        return {
            "Exhibitor": exhibitor,
            "Exhibitor Header": exhibitor_header,
            "Information": information,
            "Links": link,
            "Location": location,
            "Interests": interests,
            "Remarks": ""
        }
    except Exception as e:
        logging.error(f"Error getting company details: {str(e)}")
        return None

def create_excel_file(data, filename="mwc_exhibitors.xlsx"):
    # 데이터를 Excel 파일로 생성
    if not data:
        logging.warning("No data to write to Excel file.")
        return
    
    df = pd.DataFrame(data)
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Exhibitors')
    
    workbook = writer.book
    worksheet = writer.sheets['Exhibitors']
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    
    for column in worksheet.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        worksheet.column_dimensions[column[0].column_letter].width = max_length + 2
    
    writer.close()
    logging.info(f"Excel file '{filename}' created successfully.")

def get_chromedriver_version():
    try:
        version = subprocess.check_output(["chromedriver", "--version"]).decode("utf-8").strip()
        return version
    except subprocess.CalledProcessError as e:
        print(f"Error getting chromedriver version: {e}")
        return None

def main():
    driver = setup_driver()
    all_companies = []
    page = 1
    max_retries = 3
    
    try:
        while True:
            url = f"https://www.mwcbarcelona.com/exhibitors?page={page}"
            retries = 0
            
            while retries < max_retries:
                try:
                    driver.get(url)
                    wait_for_page_load(driver, timeout=30)  # 대기 시간을 30초로 늘림
                    
                    for i in range(1, 25):  # 24번 반복
                        xpath = f'//*[@id="traversable-list-2526362"]/ul/a[{i}]'
                        print(f"Processing XPath URL: {xpath}")
                        try:
                            link = wait_for_element(driver, xpath, by=By.XPATH)
                            if link:
                                link.click()
                                wait_for_page_load(driver, timeout=30)  # 페이지가 완전히 로드될 때까지 대기
                                
                                company_data = get_company_details(driver, driver.current_url)
                                if company_data:
                                    all_companies.append(company_data)
                                
                                # 이전 페이지로 이동
                                driver.back()
                                wait_for_page_load(driver, timeout=30)  # 페이지가 완전히 로드될 때까지 대기
                            else:
                                logging.warning(f"Link not found for XPath: {xpath}")
                        except Exception as e:
                            logging.error(f"Error processing link with XPath {xpath}: {str(e)}")
                    
                    page += 1
                    next_page_button = wait_for_element(driver, '//*[@id="item-links-2526362"]/div[3]/div/ul/li[7]/a', by=By.XPATH)
                    if next_page_button:
                        next_page_button.click()
                        wait_for_page_load(driver, timeout=30)  # 대기 시간을 30초로 늘림
                    else:
                        break
                    
                except Exception as e:
                    logging.error(f"Error on page {page}: {str(e)}")
                    retries += 1
                    if retries == max_retries:
                        logging.error(f"Failed to process page {page} after {max_retries} attempts")
                        break
                    time.sleep(5)
            
            if retries == max_retries:
                break
                
    finally:
        driver.quit()
        if all_companies:
            create_excel_file(all_companies, filename="mwc_exhibitors.xlsx")
            logging.info("Data has been saved to mwc_exhibitors.xlsx")
            logging.info(f"Total companies processed: {len(all_companies)}")
        else:
            logging.warning("No companies were processed.")

if __name__ == "__main__":
    version = get_chromedriver_version()
    if version:
        print(f"Chromedriver version: {version}")
    else:
        print("Could not determine chromedriver version.")
    main()
