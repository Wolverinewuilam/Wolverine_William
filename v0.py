import pandas as pd
import pywinauto
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import keyboard as kb
from datetime import date, timedelta
import shutil
import os
import sys
import glob
import time
import datetime
import dateutil.relativedelta

def get_password_data(file_path, file_type):
    files = [file for file in glob.glob(os.path.join(file_path, file_type)) if os.path.isfile(file)]
    return max(files, key=os.path.getctime) if files else None

def read_password_data(password_file):
    try:
        password_data = pd.read_excel(password_file, sheet_name="Query", skiprows=0)
        return (str(password_data['User'].values).strip('[]').replace("'", ""),
                str(password_data['Password'].values).strip('[]').replace("'", ""))
    except Exception as e:
        raise Exception(f"Failed to read data: {str(e)}")

def move_to_archive(folder_from, folder_to):
    print(f"move_to_archive - {datetime.datetime.now()}")
    if os.path.exists(folder_from):
        try:
            shutil.move(folder_from, folder_to, copy_function=shutil.copy2)
        except ValueError as e:
            print(e)

def copy_data(copy_from, copy_to):
    print(f"copy_data - {datetime.datetime.now()}")
    try:
        target_date = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y.%m.%d")
        target = os.path.join(copy_to, target_date)
        for file_name in os.listdir(copy_from):
            print(file_name)
            os.makedirs(target, exist_ok=True)
            shutil.copy(os.path.join(copy_from, file_name), os.path.join(target, file_name))
        print("Files are copied successfully")
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

def create_saving_path():
    os.makedirs(os.path.join(SAVING_PATH, TODAY), exist_ok=True)

def delete_specific_files():
    print(f"Del_file - {datetime.datetime.now()}")
    today_path = os.path.join(SAVING_PATH, datetime.datetime.now().strftime('%Y.%m.%d'))
    keywords = ["lawwu", "chhloan", "wsurfacepro", "dev_rfc", "dev_nco_rfc"]
    while True:
        try:
            for file in os.listdir(today_path):
                if any(keyword in file for keyword in keywords):
                    os.remove(os.path.join(today_path, file))
            print("All targeted files have been deleted.")
            break
        except OSError as e:
            print(f"Error: {e.strerror}. Retrying in 1 second...")
            time.sleep(1)

def keep_folder_by_range_of_date(date_list, path):
    try:
        for folder_name in os.listdir(path):
            if folder_name not in date_list:
                shutil.rmtree(os.path.join(path, folder_name))
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

def check_file(loading_path, saving_path, today):
    print("Checkfile")
    path = os.path.join(saving_path, today)
    create_saving_path()
    saving_files = set(os.listdir(path))
    loading_files = set(os.listdir(loading_path))
    files = list(loading_files - saving_files)
    print(f'Start - {len(loading_files)}, End - {len(files)}')
    return files

def check_excel():
    for _ in range(1200):  # 6000 seconds / 5 seconds = 1200 iterations
        try:
            Application(backend='uia').connect(title_re=".*Excel*.")
            return True
        except:
            time.sleep(5)
    return False

def close_excel():
    print("CloseExcel")
    try:
        app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.exe")
        app.window(title_re=".*Excel*.").maximize()
        app.kill()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

def open_variable(main):
    print("OpenVariable")
    main.wait("exists enabled visible ready", timeout=3600)
    while True:
        try:
            time.sleep(2)
            main.type_keys('%XYH')
            time.sleep(1)
            break
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

def login(main, user, password):
    print("Login")
    while True:
        try:
            main.set_focus()
            time.sleep(2)
            main.child_window(title="Language", auto_id="txtClient", control_type="Edit").wait("exists enabled visible ready", timeout=600)
            main.child_window(title="Language", auto_id="txtClient", control_type="Edit").set_text(CLIENT)
            time.sleep(1)
            main.child_window(auto_id="txtUser", control_type="Edit").set_text(user)
            time.sleep(1)
            main.child_window(auto_id="txtPassword", control_type="Edit").set_text(password)
            time.sleep(1)
            main.child_window(title="OK", auto_id="butOk", control_type="Button").click()
            print(f"{datetime.datetime.now()}-[get_report] login_bw successfully")
            bw_click_ok(main)
            break
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

def bw_click_ok(main):
    print(f"BW_click_OK - {datetime.datetime.now()}")
    try:
        main.child_window(title="OK", auto_id="2", control_type="Button").click()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

def file_exist_click_yes(main):
    print(f"file_exist_click_yes - {datetime.datetime.now()}")
    try:
        main.child_window(title="Yes", control_type="Button").click()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

def entry_variable(main):
    print("EntryVariable")
    main.child_window(title="PO Key Date(*)", auto_id="1001", control_type="Edit").wait("exists enabled visible ready", timeout=600)
    while True:
        try:
            main.set_focus()
            main.child_window(title="PO Key Date(*)", auto_id="1001", control_type="Edit").set_focus()
            main.child_window(title="PO Key Date(*)", auto_id="1001", control_type="Edit").set_text("")
            main.child_window(title="PO Key Date(*)", auto_id="1001", control_type="Edit").type_keys(THE_DAY_BEFORE_YESTERDAY)
            main.child_window(title="Manual Ex-Factory Date (yyyy.mm.dd)", auto_id="1001", control_type="Edit").set_focus()
            main.child_window(title="Manual Ex-Factory Date (yyyy.mm.dd)", auto_id="1001", control_type="Edit").set_text("")
            main.child_window(title="Manual Ex-Factory Date (yyyy.mm.dd)", auto_id="1001", control_type="Edit").type_keys(MANUAL_XF_DATE)
            main.child_window(title="OK", auto_id="btnOk", control_type="Button").click()
            break
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

def save_as(main):
    print("Saveas")
    main.child_window(title="File Tab", auto_id="FileTabButton", control_type="Button").wait("exists enabled visible ready active", timeout=600)
    while True:
        try:
            main.set_focus()
            main.child_window(title="File Tab", auto_id="FileTabButton", control_type="Button").set_focus()
            main.child_window(title="File Tab", auto_id="FileTabButton", control_type="Button").click()
            time.sleep(1)
            main.type_keys('%FA')
            time.sleep(1)
            main.type_keys('O')
            break
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

def close_excel_in_app(main):
    print("CloseExcelinApp")
    while True:
        try:
            main.set_focus()
            main.type_keys('%F')
            time.sleep(1)
            main.type_keys('c')
            break
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

def get_report(user, password):
    print("get_report")
    create_saving_path()
    print("Checkfile")
    path = os.path.join(SAVING_PATH, TODAY)
    delete_specific_files()
    files = check_file(LOADING_PATH, SAVING_PATH, TODAY)
    
    if not files:
        print("No files to process")
        return True

    try:
        print(f'Open Excel file - {files[0]}')
        excel = os.path.join(LOADING_PATH, files[0])
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        return False

    print("OpenExcel")
    program_path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.exe"
    app = Application(backend="uia").start(r'{} "{}"'.format(program_path, excel)).connect(path=program_path)

    if not check_excel():
        print("Excel did not open properly")
        return False

    main = app.window(title_re=".*Excel*.")

    time.sleep(2)
    try:
        main.wait('ready')
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        return False

    time.sleep(2)
    open_variable(main)
    time.sleep(2)
    print(f"{datetime.datetime.now()}-[get_report] login")
    login(main, user, password)
    time.sleep(2)
    entry_variable(main)
    time.sleep(2)
    print("Inputsaveaspath")

    file_name = os.path.join(SAVING_PATH, TODAY, files[0])
    print(file_name)

    save_as(main)
    time.sleep(2)

    while True:
        try:
            main.child_window(title="File name:", auto_id="1001", control_type="Edit").set_text("")
            time.sleep(1)
            main.child_window(title="File name:", auto_id="1001", control_type="Edit").type_keys(file_name, with_spaces=True)
            time.sleep(1)
            main.child_window(title="Save", auto_id="1", control_type="Button").click()
            print(f"{datetime.datetime.now()}-[get_report] File Save - {files[0]}")
            break
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

    print("Check Excel exist or not")
    
    for _ in range(1200):  # 6000 seconds  / 5 seconds = 1200 iterations
        if os.path.exists(file_name):
            break
        time.sleep(15)
    else:
        print("File did not save within the expected time")
        return False

    time.sleep(1)
    close_excel_in_app(main)
    time.sleep(1)
    close_excel()
    return True

if __name__ == "__main__":
    while True:        
        TODAY = datetime.datetime.now().strftime('%Y.%m.%d')
        THE_DAY_BEFORE_YESTERDAY = (datetime.datetime.now() - datetime.timedelta(days=2)).strftime('%m/%d/%Y')
        LOADING_PATH = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Documents\Query\PO Detail - New Dates -7.0"
        SAVING_PATH = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Data - Power BI\PO Detail - New Dates -7.0"
        # SAVING_PATH = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Test\PDF_TO_EXCEL"
        BACKUP_PATH = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Power Query and Dax Training\Data Keep\PO Detail - New Dates -7.0"
        
        MANUAL_XF_DATE =   "1/1/2022-12/31/2028"
        CLIENT = "010"
        password_location = r"D:\User\lawwu\OneDrive - Wolverine World Wide\Documents\Query\Password"
        file_type = "Password.xlsx"
        password_data = get_password_data(password_location, file_type)
        user, password = read_password_data(password_data)
        today = datetime.datetime.now()
        yesterday_path = os.path.join(SAVING_PATH, (today - datetime.timedelta(days=1)).strftime('%Y.%m.%d'))
        archive_path = os.path.join(SAVING_PATH, "Archive")
        copy_data(yesterday_path, BACKUP_PATH)
        move_to_archive(yesterday_path, archive_path)
        _3day_path = os.path.join(SAVING_PATH, (today - datetime.timedelta(days=3)).strftime('%Y.%m.%d'))
        move_to_archive(_3day_path, archive_path)
        date_list = pd.date_range((today - dateutil.relativedelta.relativedelta(days=13)).date(), today.date(), freq='D').strftime('%Y.%m.%d')
        keep_folder_by_range_of_date(date_list, archive_path)
        close_excel()
        
        files_to_process = check_file(LOADING_PATH, SAVING_PATH, TODAY)
        if not files_to_process:
            print("No files to process. Stopping the run.")
            break
        
        try:
            if get_report(user, password):
                files_to_process = check_file(LOADING_PATH, SAVING_PATH, TODAY)
                if not files_to_process:
                    print("All files processed. Stopping the run.")
                    break
            else:
                print("Failed to process all files. Retrying...")
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
        
        delete_specific_files()
        print("Completed one iteration. Checking for more files to process...")

    print("Script execution completed.")