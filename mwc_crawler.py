import subprocess
import sys

# 필요한 패키지가 설치되어 있는지 확인
try:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium", "webdriver_manager", "pandas", "openpyxl", "tqdm", "keyboard"])
except subprocess.CalledProcessError as e:
    print(f"Error installing packages: {e}")
    sys.exit(1)

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

def get_company_details(driver, company_url):
    # 회사 세부 정보를 가져옴
    try:
        driver.get(company_url)
        
        exhibitor = wait_for_element(driver, "h1.text-3xl.font-medium.leading-none.lg:text-5xl.font-black.font-heading")
        exhibitor = exhibitor.text.strip() if exhibitor else ""
        
        information = wait_for_element(driver, "#maincontent > div")
        information = information.text.strip() if information else ""
        
        location = ""
        interests = ""
        
        aside_elements = driver.find_elements(By.CSS_SELECTOR, "#exhibitor-container > aside > div")
        for element in aside_elements:
            heading = element.find_element(By.CSS_SELECTOR, "h2").text.strip()
            if heading == "Location":
                location_elements = element.find_elements(By.CSS_SELECTOR, "ul li")
                location = ", ".join([loc.text.strip() for loc in location_elements])
            elif heading == "Interests":
                interests_elements = element.find_elements(By.CSS_SELECTOR, "ul li")
                interests = ", ".join([int.text.strip() for int in interests_elements])

        return {
            "Exhibitors": exhibitor,
            "Location": location,
            "Interests": interests,
            "Information": information,
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
                    company_links = driver.find_elements(By.CSS_SELECTOR, "div.exhibitor-card a")
                    if not company_links:
                        if page > 1:
                            logging.info("Reached the last page.")
                            break
                        else:
                            retries += 1
                            continue
                    
                    logging.info(f"Processing page {page} ({len(company_links)} companies)")
                    
                    for link in tqdm(company_links, desc=f"Page {page}"):
                        company_url = link.get_attribute('href')
                        driver.get(company_url)  # 회사의 세부 페이지로 이동
                        
                        # 스페이스바 입력 대기
                        print("Press the spacebar to continue...")
                        keyboard.wait('space')
                        
                        company_data = get_company_details(driver, company_url)
                        if company_data:
                            all_companies.append(company_data)
                        
                        # 이전 페이지로 이동
                        driver.back()
                        wait_for_page_load(driver, timeout=30)  # 대기 시간을 30초로 늘림
                        
                        # 스페이스바 입력 대기
                        print("Press the spacebar to continue to the next company...")
                        keyboard.wait('space')
                    
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
    main()
