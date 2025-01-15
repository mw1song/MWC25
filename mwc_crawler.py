import subprocess
import sys
import signal
import warnings  # 경고 무시를 위한 모듈 추가
import os  # 파일 경로 처리를 위한 모듈 추가

# 필요한 패키지가 설치되어 있는지 확인
# try:
#     subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium", "webdriver_manager", "pandas", "openpyxl", "tqdm", "keyboard"])
# except subprocess.CalledProcessError as e:
#     print(f"Error installing packages: {e}")
#     sys.exit(1)

# 경고 메시지 무시 설정
warnings.filterwarnings("ignore", category=DeprecationWarning)

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

def accept_cookies(driver):
    try:
        accept_button = wait_for_element(driver, 'button.accept-cookies', timeout=10)
        if accept_button:
            accept_button.click()
            logging.info("Accepted cookies.")
        else:
            logging.info("No cookies acceptance button found.")
    except Exception as e:
        logging.error(f"Error accepting cookies: {str(e)}")

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
        return element.text.strip() if element else "N/A"
    except NoSuchElementException:
        return "N/A"

def get_company_details(driver, company_url):
    try:
        driver.get(company_url)
        print(f"Opened URL: {company_url}")  # 새로운 URL 출력
        
        exhibitor = get_text_or_null(driver, '//*[@id="headerContainer"]/div/div[3]/nav/ul/li[5]')
        print(f"Exhibitor: {exhibitor}")
        
        exhibitor_header = [
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[1]/span'),
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[2]/span'),
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[3]/span'),
            get_text_or_null(driver, '//*[@id="exhibitor-header"]/div/div[2]/div[1]/span[4]/span')
        ]
        
        information = get_text_or_null(driver, '//*[@id="maincontent"]/div')
        
        links = ["N/A"] * 6
        locations = ["N/A"] * 6
        interests = ["N/A"] * 6
        
        aside_elements = driver.find_elements(By.XPATH, '//*[@id="exhibitor-container"]/aside/div')
        for i, element in enumerate(aside_elements, start=1):
            heading = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{i}]/h5')
            print(f"Heading: {heading}")
            if heading == "CONTACT & LINKS":
                for j in range(1, 7):
                    links[j-1] = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{1}]/ul/li[{j}]/a')
            elif heading == "LOCATION":                             
                for j in range(1, 7):
                    locations[j-1] = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{2}]/ul/p[{j}]')
            elif heading == "INTERESTS":
                for j in range(1, 7):
                    interests[j-1] = get_text_or_null(driver, f'//*[@id="exhibitor-container"]/aside/div[{3}]/ul/li[{j}]')
        
        return {
            "Exhibitor": exhibitor,
            "Exhibitor Header": exhibitor_header,
            "Information": information,
            "Links": links,
            "Location": locations,
            "Interests": interests
        }
    except Exception as e:
        logging.error(f"Error getting company details: {str(e)}")
        return None

def process_xpath_url(driver, xpath):
    try:
        link = wait_for_element(driver, xpath, by=By.XPATH)
        if link:
            link.click()
            wait_for_page_load(driver, timeout=30)  # 페이지가 완전히 로드될 때까지 대기
        
            print(f"Current URL: {driver.current_url}")
            company_data = get_company_details(driver, driver.current_url)
            if company_data:
                all_companies.append(company_data)
        else:
            logging.warning(f"Link not found for XPath: {xpath}")
    except Exception as e:
        logging.error(f"Error processing link with XPath {xpath}: {str(e)}")

def create_excel_file(data, filename="mwc_exhibitors.xlsx"):
    # 데이터를 Excel 파일로 생성
    if not data:
        logging.warning("No data to write to Excel file.")
        return
    
    # 절대 경로로 파일 경로 설정
    filepath = os.path.abspath(filename)
    
    # 데이터 변환
    transformed_data = []
    for item in data:
        base_data = {
            "Exhibitor": item["Exhibitor"],
            "Exhibitor Header 1": item["Exhibitor Header"][0],
            "Exhibitor Header 2": item["Exhibitor Header"][1],
            "Exhibitor Header 3": item["Exhibitor Header"][2],
            "Exhibitor Header 4": item["Exhibitor Header"][3],
            "Information": item["Information"]
        }
        # Links
        for idx in range(6):
            base_data[f"Link {idx + 1}"] = item["Links"][idx]
        # Location
        for idx in range(6):
            base_data[f"Location {idx + 1}"] = item["Location"][idx]
        # Interests
        for idx in range(6):
            base_data[f"Interest {idx + 1}"] = item["Interests"][idx]
        
        transformed_data.append(base_data)
    
    df = pd.DataFrame(transformed_data)
    writer = pd.ExcelWriter(filepath, engine='openpyxl')
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
    logging.info(f"Excel file '{filepath}' created successfully.")

def handle_exit(signum, frame):
    logging.info("Execution interrupted. Saving progress...")
    create_excel_file(all_companies, filename="mwc_exhibitors.xlsx")
    logging.info("Progress saved. Exiting...")
    driver.quit()  # 드라이버 종료
    sys.exit(0)

def handle_popup(driver):
    try:
        popup_button = wait_for_element(driver, '//*[@id="onetrust-accept-btn-handler"]', by=By.XPATH, timeout=10)
        if popup_button:
            popup_button.click()
            logging.info("Popup accepted.")
        else:
            logging.info("No popup button found.")
    except Exception as e:
        logging.error(f"Error handling popup: {str(e)}")

def main():
    global all_companies, driver
    driver = setup_driver()
    all_companies = []
    page = 1
    max_retries = 3
    max_pages = 2  # 설정된 페이지 수만큼 반복

    # 강제 종료 시 처리
    signal.signal(signal.SIGINT, handle_exit)
    signal.signal(signal.SIGTERM, handle_exit)
    
    try:
        while page <= max_pages:
            print(f"page: {page}")
            url = f"https://www.mwcbarcelona.com/exhibitors?page={page}"
            print(f"Processing URL: {url}")
            retries = 0
                        
            while retries < max_retries:
                try:
                    driver.get(url)
                    wait_for_page_load(driver, timeout=30)  # 대기 시간을 30초로 늘림
                    accept_cookies(driver)  # 쿠키 수락
                    handle_popup(driver)  # 팝업 처리
                    
                    for i in range(2, 5):  # 24번 반복
                        xpath = f'//*[@id="traversable-list-2526362"]/ul/a[{i}]'
                        print(f"Processing XPath URL: {xpath}")
                        process_xpath_url(driver, xpath)
                    
                    page += 1
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

