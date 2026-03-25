import datetime
import glob
import os
import shutil
from shutil import copy2
import sys
import time
import pandas as pd
from pywinauto import Application
from selenium import webdriver
from chromedriver_py import binary_path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
    
def switch_season(Season):
    season = Season[0]+Season[-2:]
    return season

def switch_brand(Brand):
    brand = {
        "MERRELL": "Merrell",
        "SAUCONY": "Saucony",
        "SPERRY TOP-SIDER": "Sperry",
        "CATERPILLAR FOOTWEAR": "CAT",
        "KEDS": "Keds",
        "WOLVERINE CHILDREN'S GROUP": "Kids Group",
        "WOLVERINE": "Wolverine",
        "HUSH PUPPIES": "Hush Puppies",
        "CHACO": "Chaco",
        "BATES": "Bates",
        "HARLEY DAVIDSON FOOT": "Harley",
        "HY-TEST": "Hy-test"
    }
    return brand.get(Brand, "")

def switch_seasons_brands(Brands):
    brands = {
        "Spring/Summer 2027": ["BATES", "CATERPILLAR FOOTWEAR", "HARLEY DAVIDSON FOOT", "HY-TEST", "WOLVERINE", "MERRELL"],
        "Fall/Winter 2027": ["BATES", "CATERPILLAR FOOTWEAR", "HARLEY DAVIDSON FOOT", "HY-TEST", "WOLVERINE", "MERRELL"]
    }
    return brands.get(Brands, "Unknown Season")

def create_saving_path(download_path, time_to_wait):
    time_counter = 0
    while time_counter <= time_to_wait:
        print(f"create_saving_path - {datetime.datetime.now()}")
        if os.path.exists(download_path):
            break
        os.makedirs(download_path)
        print(f"create_saving_path successful - {datetime.datetime.now()}")
        time_counter += 1
        time.sleep(1)

def PLM_login(download_path, URL, user, password, time_to_wait, all_season):
    print(f"PLM_login - {datetime.datetime.now()}")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs",
                                {"download.default_directory": download_path,
                                    "download.prompt_for_download": True
                                }
                                )
    try:
        print(f"Using chromedriver_py")
        svc = webdriver.ChromeService(executable_path=binary_path)
        driver = webdriver.Chrome(service=svc, options=options)
    except:
        try:
            print(f"Using ChromeDriverManager")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        except:
            print(f"Using chromedriver")
            binary_path = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Documents\Query\ChromeDriver\chromedriver-win64\chromedriver.exe"
            service_object = Service(binary_path)
            driver = webdriver.Chrome(service=service_object, options=options)
    driver.implicitly_wait(300)
    driver.maximize_window()
    driver.get(URL)
    driver.find_element(By.NAME, "username").click()
    driver.find_element(By.NAME, "username").send_keys(user)
    driver.find_element(By.NAME, "password").click()
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.XPATH, "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/div[5]/button[1]").click()
    print(f"PLM_login successful - {datetime.datetime.now()}")
    time.sleep(1)
    create_saving_path(download_path, time_to_wait)
    loop_for_season_brand(all_season, download_path, driver, time_to_wait)

def click_search(driver, time_to_wait):
    item = "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/ul[1]/li[2]/a[1]"
    time_counter = 0
    checkstatus = False
    while not checkstatus:
        try:
            print(f"click_search - {datetime.datetime.now()}")
            driver.find_element(By.XPATH, item).click()
            checkstatus = True
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:break

def click_plm_export(driver, time_to_wait):
    item = "//div[contains(text(),'All')]" #Remember to save search as "All" in PLM (You can rename to your bookmark)
    time_counter = 0
    checkstatus = False
    while not checkstatus:
        try:
            print(f"click_plm_export - {datetime.datetime.now()}")
            driver.find_element(By.XPATH, item).click()
            checkstatus = True
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:break

def select_brand_season(driver, brand, season):
    print(f"select_{brand}_and_{season} - {datetime.datetime.now()}")
    brand_drop=Select(driver.find_element(By.XPATH, "/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[3]/form[4]/div[2]/div[3]/span[1]/select[1]"))
    season_drop =Select(driver.find_element(By.XPATH, "/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[3]/form[4]/div[2]/div[4]/span[1]/select[1]"))
    brand_drop.select_by_visible_text(brand)
    season_drop.select_by_visible_text(season)

def click_export_search(driver, time_to_wait):
    item = "/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[8]/div[15]/a[1]"
    time_counter = 0
    checkstatus = False
    while not checkstatus:
        try:
            print(f"click_export_search - {datetime.datetime.now()}")
            driver.find_element(By.XPATH, item).click()
            checkstatus = True
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:break

def click_export(driver, time_to_wait):
    item = "/html[1]/body[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[8]/div[10]/a[1]"
    time_counter = 0
    checkstatus = False
    while not checkstatus:
        try:
            print(f"click_export - {datetime.datetime.now()}")        
            driver.find_element(By.XPATH, item).click()
            checkstatus = True
            time.sleep(2)
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:break

def CheckChrome(time_to_wait, FileName):
    print(f"CheckChrome - {datetime.datetime.now()}")
    time_counter = 0
    checkstatus = False    
    while not checkstatus:
        try:
            Application(backend ='uia').connect(title_re ='.*Untitled - Google Chrome*.').is_process_running()
            print("Chrome is processing")
            checkstatus = True
            download_file(time_to_wait, FileName)
        except:
            time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:break

def download_file(time_to_wait, FileName):
    print(f"download_file - {datetime.datetime.now()}")
    app = Application(backend='uia').connect(title_re='.*Untitled - Google Chrome*.')
    main = app.window(title_re='.*Untitled - Google Chrome*.')
    if not os.path.exists(os.path.dirname(os.path.abspath(FileName))):
        os.makedirs(os.path.dirname(os.path.abspath(FileName)))

    if _save_file_with_retries(main, FileName, time_to_wait):
        print(f"File successfully saved: {FileName}")
    else:
        print(f"Failed to save file: {FileName}")

def _set_filename_and_save(main, FileName):
    file_input = main.child_window(title="File name:", auto_id="1001", control_type="Edit")
    file_input.set_text("")
    file_input.type_keys(FileName, with_spaces=True)
    current_value = file_input.get_value()
    while current_value != FileName:
        file_input.set_text("")
        file_input.type_keys(FileName, with_spaces=True)
        current_value = file_input.get_value()
    main.child_window(title="Save", auto_id="1", control_type="Button").click()
def _wait_for_file(FileName, timeout):
    for _ in range(timeout):
        if os.path.exists(FileName):
            return True
        time.sleep(1)
    print(f"Timeout waiting for file: {FileName}")
    return False

def _save_file_with_retries(main, FileName, time_to_wait, max_retries=5):
    for attempt in range(max_retries):
        try:
            _set_filename_and_save(main, FileName)
            print(f"File Save as - {datetime.datetime.now()} {'='*50} \n{FileName}")
            try:
                main.child_window(title="OK", auto_id="CommandButton_1", control_type="Button").click()
                print(f"Error dialog detected and handled (attempt {attempt + 1})")
                time.sleep(1)
                continue
            except:
                pass
            return _wait_for_file(FileName, time_to_wait)
        except Exception as e:
            print(f"Error on attempt {attempt + 1}: {e}")
            time.sleep(1)    
    return False

def loop_for_season_brand(all_season, download_path, driver, time_to_wait):
    for season in all_season: 
        for brand in reversed(switch_seasons_brands(season)):
            print(f"loop_for_season_brand - {datetime.datetime.now()}")
            FileName = f'{download_path}\\PLM Export - {switch_brand(brand)} - {switch_season(season)}.xlsx'
            print(FileName)            
            if not os.path.exists(FileName):
                time.sleep(1)
                driver.get("https://plm.wwwinc.com/Search")
                click_search(driver, time_to_wait)
                click_plm_export(driver, time_to_wait)
                select_brand_season(driver, brand, season)
                click_export_search(driver, time_to_wait)
                click_export(driver, time_to_wait)
                CheckChrome(time_to_wait, FileName)
            else:
                next

def main(all_season, download_path, time_to_wait):
    user = "harris.le_ic@wwwinc.com"
    password = "xxx"
    URL = f"https://plm.wwwinc.com/login"
    PLM_login(download_path, URL, user, password, time_to_wait, all_season)
    
if __name__ == "__main__":
    save_path = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Data - Power BI\CAP\PLM Export"
    today = datetime.datetime.now()
    download_path = os.path.join(save_path, today.strftime('%Y.%m.%d'))
    time_to_wait = 12000
    create_saving_path(download_path, time_to_wait)
    all_season=["Spring/Summer 2027", "Fall/Winter 2027"]
    checkstatus = False
    try:
        while not checkstatus:
            Filescheck = []
            Files = os.listdir(download_path)
            for filecheck in Files:
                if any(keyword in filecheck for keyword in map(switch_season, all_season)):
                    Filescheck.append(filecheck)
                    print(f"Files found: {len(Filescheck)}")
            Target = 0
            for season in all_season:
                Target += len(switch_seasons_brands(season))
            print(f"Target files: {Target}")
            print(f"Current files: {len(Filescheck)}")
            print(f"Remaining: {Target - len(Filescheck)}")
            
            if not len(Filescheck) == Target:
                print("Starting download process...")
                main(all_season, download_path, time_to_wait)
            else:
                print("All files downloaded successfully!")
                checkstatus = True
                time.sleep(1)
        time.sleep(5)
    except KeyboardInterrupt:
        print("Process interrupted by user.")

    print("Completed PLM Export!")