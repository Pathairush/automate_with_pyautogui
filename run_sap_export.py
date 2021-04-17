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

def check_server_location(sap_server, location1, location2):

    '''
    This function check that the provided SAP server exists in the destop
    '''

    try:
        server = pyautogui.locateOnScreen(location1)
        if server is None:
            server = pyautogui.locateOnScreen(location2)
            if server is None:
                raise ValueError(f"Cannot detect {sap_server} sap server location")
    except Exception as e:
        print(e)
        raise 

    return server

def remove_export_file_if_exists():

    '''
    this function remove an export file if exists, the reason is to prevent the conflict in export file saving.
    '''

    try:
        if os.path.isfile(os.path.join(config['PATH']['LOCAL'], config['FILE']['EXPORT'])):
            os.remove(os.path.join(config['PATH']['LOCAL'], config['FILE']['EXPORT']))
            print(f'-- Found {config["FILE"]["EXPORT"]} file exists remove it first.')
    except Exception as e:
        print(e)
        raise
        
def open_sap_gui(config):
    
    '''
    this function open the SAP GUI application in your desktop, you can change your saplogon.exe path in configuration.
    '''

    try: 
        sap_gui = subprocess.Popen(config['PATH']['SAP_LOGON'])
        print('-- Open SAP GUI')
        time.sleep(5)
        
        return sap_gui

    except Exception as e:
        pritn(e)
        raise

def select_sap_server(sap_server, server_location, server_location_default):

    '''
    this function select the sap server based on the provided server location reference.
    '''
    try: 
        server = check_server_location(sap_server, server_location, server_location_default)
        while server is None:
            time.sleep(3)
            server = check_server_location(sap_server, server_location, server_location_default)
        pyautogui.moveTo(server)
        pyautogui.doubleClick()
        print(f'-- Open SAP {sap_server} server')
        time.sleep(10)
    
    except Exception as e:
        print(e)
        raise

def fill_username_password(username, password):

    '''
    this function fill username and password to the SAP logon application.
    '''

    try:
        pyautogui.typewrite(username, interval=0.1)
        pyautogui.press('tab')
        pyautogui.typewrite(password, interval=0.1)
        pyautogui.press('enter')
        print(f'-- SAP LOGGED IN')
        time.sleep(3)
    except Exception as e:
        print(e)
        raise
    
def open_zfiaraging_module():

    '''
    this function loads the ZFIARAGIN module from the SAP main screen.
    '''
    try:
        pyautogui.typewrite('ZFIARAGING')
        pyautogui.press('enter')
        print('-- Open ZFIARAGING module')
        time.sleep(3)
    except Exception as e:
        print(e)
        raise e

def fill_zfiaraging_form(comp_code, profit_center, aging_date, num_tab_aging_date):

    '''
    this function fill the information from input parameters to the ZFIARAGING program and execute it.
    '''

    try:
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
        pyautogui.press('down') # select Display ALV Grid new Standard
        pyautogui.press('f8')
        time.sleep(10)
    except Exception as e:
        print(e)
        raise
    
def check_report_load():

    '''
    check that the ZFIARAGING has been loaded before moving to the next process.
    '''

    try:
        check_report_load = pyautogui.locateOnScreen('locations/report_load_success.png')
        retry = 0
        while not check_report_load:
            time.sleep(30)
            check_report_load = pyautogui.locateOnScreen('locations/report_load_success.png')
            retry += 1
            if retry == 6:
                raise ValueError("-- Report load failed : Report cannot be loaded after 180 seconds.")

        print('-- Report load success')
    except Exception as e:
        print(e)
        raise

def select_report_layout(report_layout):

    '''
    select exported layout for SAP report
    '''

    try:
        pyautogui.hotkey('ctrl', 'f9') # shortcut to call repory layout option.
        time.sleep(10)
        pyautogui.moveTo(report_layout)
        pyautogui.click()
        time.sleep(10)
    except Exception as e:
        print(e)
        raise

def process_exists(process_name):

    '''
    this function check that the provided process name is existed, if exists return true, else false.
    '''

    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    output = subprocess.check_output(call).decode()
    last_line = output.strip().split('\r\n')[-1]
    
    return last_line.lower().startswith(process_name.lower())

def export_report():

    '''
    export report to excel spraed sheet.
    '''
    
    try:
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
                raise ValueError('-- FAILED : cannot save report to export.xlsx after 60 seconds.')
    
    except Exception as e:
        print(e)
        raise

def kill_process(process_name):

    '''
    kill window process if exists.
    '''

    try:
        os.system(f'TASKKILL /F /IM {process_name}')
        time.sleep(10)
    except Exception as e:
        print(e)
        raise
            
def rename_export_file(config, export_date, key, profit_center, sap_server):

    '''
    rename and move export file to target directory for uploading to S3.
    '''

    export_file_name = os.path.join(config['PATH']['LOCAL'], config['FILE']['EXPORT'])
    target_dir = os.path.join(config['PATH']['LOCAL'], export_date)

    file_name = '{}_{}_{}.xlsx'.format(key, profit_center, sap_server)
    full_path_file = os.path.join(config['PATH']['LOCAL'], export_date, file_name)
    
    try:
        # check if export file exists
        if os.path.isfile(export_file_name): 

            # check export_aging_date folder exists
            if not os.path.exists(target_dir): 
                print(f"{target_dir} doesn't exists ; create a new folder")
                os.mkdir( 
                    os.path.join(target_dir) 
                )

            # if target file exists, remove the exist file and replace with current one
            if os.path.isfile(full_path_file): 
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
    except Exception as e:
        print(e)
        raise

def extract_zfiaraging_report(key, comp_code, profit_center, aging_date, export_date, config, sap_server):

    '''
    running SAP Report Extraction program.
    '''

    if sap_server == 'dev':

        server_location = config["LOCATION"]["SAP_DEV"]
        server_location_default = config["LOCATION"]["SAP_DEV_DEFAULT"]
        username = config['CREDENTIALS']['SAP_DEV_USERNAME']
        password = config['CREDENTIALS']['SAP_DEV_PASSWORD']
        num_tab_aging_date = 18
        report_layout = config['LOCATION']['LAYOUT_DEV']

    elif sap_server == 'prod':

        server_location = config["LOCATION"]["SAP_PRD"]
        server_location_default = config["LOCATION"]["SAP_PRD"]
        username = config["CREDENTIALS"]["SAP_PRD_USERNAME"]
        password = config["CREDENTIALS"]["SAP_PRD_PASSWORD"]
        num_tab_aging_date = 15
        report_layout = config["LOCATION"]["LAYOUT_PRD"] 

    else:
        raise ValueError(f"There is no {sap_server} choice, please check either ['dev', 'prod']")

    try:
        remove_export_file_if_exists()
        sap_gui = open_sap_gui(config)
        select_sap_server(sap_server, server_location, server_location_default)
        fill_username_password(username, password)
        open_zfiaraging_module()
        fill_zfiaraging_form(comp_code, profit_center, aging_date, num_tab_aging_date)
        check_report_load()
        select_report_layout(report_layout) 
        export_report()
        kill_process('excel.exe')
        rename_export_file(config, export_date, key, profit_center, sap_server)
        sap_gui.terminate()
    except Exception as e:
        print(e)
        raise

def main():

    # Receive input parameters from command line.
    parser = argparse.ArgumentParser()
    parser.add_argument('--aging_date')
    parser.add_argument('--sap_server')
    args = parser.parse_args()

    # assign arguments
    aging_date = args.aging_date if args.aging_date is not None else (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%d.%m.%Y")
    export_date = aging_date.replace('.', '')
    sap_server = args.sap_server if args.sap_server is not None else "dev"

    print(f'-- SAP SERVER : "{sap_server.upper()}" -- AGING DATE :  {aging_date} -- EXPORT DATE : {export_date}')

    _input = pd.read_csv('../CONFIG_compcode_subcompname.csv', dtype={'profit_center':str})
    _input = _input[['sub_comp_name', 'comp_code', 'profit_center']]
    for params in _input.itertuples():

        try:
            # extract input parameters
            _, key, comp_code, profit_center = params
            key, comp_code, profit_center = str(key), str(comp_code), str(profit_center)
            
            print(f'-- {datetime.datetime.now()} - extract report for {key}')
            extract_zfiaraging_report(key, comp_code, profit_center, aging_date, export_date, config, sap_server)
        except Exception as e:
            print(e)
            continue

if __name__ == "__main__":
    
    main()