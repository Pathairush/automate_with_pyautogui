import subprocess
import time
import os
import pandas as pd
import numpy as np
import datetime

import pyautogui
pyautogui.PAUSE = 0.3
pyautogui.FAILSAFE = True

import configparser
config = configparser.ConfigParser()
config.read('config.cfg')

import argparse

def close_excel():
    try:
        os.system('TASKKILL /F /IM excel.exe')
        time.sleep(10)
    except Exception:
        print("KU")

def process_exists(process_name):

    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    output = subprocess.check_output(call).decode()
    # check in last line for process name
    last_line = output.strip().split('\r\n')[-1]
    # because Fail message could be translated
    return last_line.lower().startswith(process_name.lower())

def check_server_location(SAP_SERVER):
    if SAP_SERVER == "DEV":
        
        location1 = config["LOCATION"]["SAP_DEV"]
        location2 = config["LOCATION"]["SAP_DEV_DEFAULT"]
    
    if SAP_SERVER == 'PRD':

        location1 = config["LOCATION"]["SAP_PRD"]
        location2 = config["LOCATION"]["SAP_PRD"]
    
    server = pyautogui.locateOnScreen(location1)
    if server is None:
        server = pyautogui.locateOnScreen(location2)
        if server is None:
            raise ValueError(f"Cannot detect {SAP_SERVER} sap server location")
    return server

def remove_export_file_if_exists():
    if os.path.isfile(os.path.join(config['PATH']['LOCAL'], config['FILE']['EXPORT'])):
        os.remove(os.path.join(config['PATH']['LOCAL'], config['FILE']['EXPORT']))
        print(f'-- Found {config["FILE"]["EXPORT"]} file exists remove it first.')
        
def open_sap_gui(config):
    sap_gui = subprocess.Popen(config['PATH']['SAP_LOGON'])
    print('-- Open SAP GUI')
    time.sleep(5)
    
    return sap_gui

def select_sap_server(SAP_SERVER):
    
    server = check_server_location(SAP_SERVER)
    while server is None:
        time.sleep(3)
        server = check_server_location(SAP_SERVER)
    pyautogui.moveTo(server)
    pyautogui.doubleClick()
    print(f'-- Open SAP {SAP_SERVER} server')
    time.sleep(10)
    
    
def fill_username_password(SAP_SERVER):

    if SAP_SERVER == "DEV":
        
        username = config['CREDENTIALS']['SAP_DEV_USERNAME']
        password = config['CREDENTIALS']['SAP_DEV_PASSWORD']
    
    if SAP_SERVER == 'PRD':

        username = config["CREDENTIALS"]["SAP_PRD_USERNAME"]
        password = config["CREDENTIALS"]["SAP_PRD_PASSWORD"]

    pyautogui.typewrite(username, interval=0.1)
    pyautogui.press('tab')
    pyautogui.typewrite(password, interval=0.1)
    pyautogui.press('enter')
    print(f'-- SAP {SAP_SERVER} LOGGED IN')
    time.sleep(3)
    
def open_zfiaraging_module():
    pyautogui.typewrite('ZFIARAGING')
    pyautogui.press('enter')
    print('-- Open ZFIARAGING module')
    time.sleep(3)

def fill_zfiaraging_form(comp_code, profit_center, aging_date, SAP_SERVER):

    if SAP_SERVER == "DEV":
        
        num_tab_aging_date = 18
    
    if SAP_SERVER == 'PRD':

        num_tab_aging_date = 15

    pyautogui.typewrite(comp_code)
    for i in range(0,3):
        pyautogui.press('tab')

    if str(profit_center) == "nan":
        pass
    else:
        pyautogui.typewrite(profit_center)

    for i in range(0,num_tab_aging_date):
        pyautogui.press('tab')
    pyautogui.typewrite(aging_date)
    pyautogui.press('tab')
    pyautogui.press('down') # select Display ALV Grid Standard
    pyautogui.press('f8')
    time.sleep(10)
    
def check_report_load():
    check_report_load = pyautogui.locateOnScreen('locations/report_load_success.png')
    retry = 0
    while not check_report_load:
        time.sleep(30)
        check_report_load = pyautogui.locateOnScreen('locations/report_load_success.png')
        retry += 1
        if retry == 6:
            print('-- Report load failed')
            break
    print('-- Report load success')
    
def select_report_layout(SAP_SERVER):

    if SAP_SERVER == "DEV":
        
        report_layout = config['LOCATION']['LAYOUT_DEV']
    
    if SAP_SERVER == 'PRD':

        report_layout = config["LOCATION"]["LAYOUT_PRD"]

     # to select report layout
    pyautogui.hotkey('ctrl', 'f9')
    time.sleep(10)
    pyautogui.moveTo(report_layout)
    pyautogui.click()
    time.sleep(10) # wait for new layout to reload
    
def export_report():
    
    # move to center of screen
    pyautogui.moveTo(pyautogui.size().width / 2,
                    pyautogui.size().height / 2)
    pyautogui.click(button='right')
    time.sleep(5)
    pyautogui.moveRel(10,150) # move to spreadsheet buttong
    pyautogui.click()
    time.sleep(5)
    pyautogui.press('enter') # alredy set default export to XLSX
    time.sleep(15)
    print('-- Save report to export.xlsx')
    pyautogui.press('enter') # to save report to export.xlsx
    time.sleep(10)
    
    retry = 0
    while not process_exists('excel.exe'):
        pyautogui.press('enter') # to save report to export.xlsx
        time.sleep(10)
        retry += 1
        if retry == 6:
            print('retry exceed')
            break
            
def rename_export_file(config, export_aging_date, key, profit_center, SAP_SERVER):

    file_name = '{}_{}_{}.xlsx'.format(key, profit_center, SAP_SERVER)
    full_path_file = os.path.join(config['PATH']['LOCAL'], export_aging_date, file_name)
    
    if os.path.isfile(os.path.join(config['PATH']['LOCAL'], config['FILE']['EXPORT'])): # check if export file exists

        if not os.path.exists(os.path.join(config['PATH']['LOCAL'], export_aging_date)): # check export_aging_date folder exists
            print(f"{export_aging_date} folder doesn't exists ; create a new folder")
            os.mkdir( 
                os.path.join(config['PATH']['LOCAL'], export_aging_date) 
            )

        if os.path.isfile(full_path_file): # if target file exists
            # remove the exist file and replace with current one
            print(f"{file_name} is exists ; remove the exist file and replace with current one")
            os.remove(full_path_file)

        # move export file to target folder and rename
        os.rename(
            os.path.join(config['PATH']['LOCAL'], f"export.XLSX"),
            full_path_file
                 )
        print(f'''
            Rename export file to {full_path_file}
        ''')

def extract_zfiaraging_report(key, comp_code, profit_center, aging_date, export_aging_date, config, SAP_SERVER):
    
    remove_export_file_if_exists()
    sap_gui = open_sap_gui(config)
    select_sap_server(SAP_SERVER)
    fill_username_password(SAP_SERVER)
    open_zfiaraging_module()
    fill_zfiaraging_form(comp_code, profit_center, aging_date, SAP_SERVER)
    check_report_load()
    select_report_layout(SAP_SERVER) 
    export_report()
    close_excel()
    rename_export_file(config, export_aging_date, key, profit_center, SAP_SERVER)
    sap_gui.terminate()


def main():

    # start time 9.21 PM, end time 10.14 PM (~ 53 mins per round on SAP DEV env)
    parser = argparse.ArgumentParser()
    parser.add_argument('--aging_date')
    args = parser.parse_args()

    _input = pd.read_csv('../CONFIG_compcode_subcompname.csv', dtype={'profit_center':str})
    _input = _input[['sub_comp_name', 'comp_code', 'profit_center']]

    aging_date = args.aging_date if args.aging_date is not None else (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%d.%m.%Y")
    export_aging_date = aging_date.replace('.', '')
    SAP_SERVER = "PRD"

    print(f'-- RUNNING ON {SAP_SERVER} ENVIRONMENT TO EXTRACT REPORT AS OF AGING DATE {aging_date}, {export_aging_date}')

    for _tup in _input.itertuples():

        _, key, comp_code, profit_center = _tup
        if profit_center == '2472501000':
            key = str(key)
            comp_code = str(comp_code)
            profit_center = str(profit_center)
            print(f'-- {datetime.datetime.now()} - extract report for {key}')
            extract_zfiaraging_report(key, comp_code, profit_center, aging_date, export_aging_date, config, SAP_SERVER)

if __name__ == "__main__":
    
    main()