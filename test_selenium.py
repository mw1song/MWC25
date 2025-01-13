# filepath: /d:/Coding/MWC25/test_selenium.py
from selenium import webdriver

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)
    return driver

def main():
    driver = setup_driver()
    driver.get("https://www.google.com")
    print(driver.title)
    driver.quit()

if __name__ == "__main__":
    main()