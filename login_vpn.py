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

def main():

    subprocess.Popen("C:\Program Files\Fortinet\FortiClient\FortiClient.exe")
    time.sleep(5)

    # move to center
    pyautogui.moveTo(pyautogui.size().width / 2,
                    pyautogui.size().height / 2)
    pyautogui.moveRel(0,170)                         
    pyautogui.click()
    pyautogui.typewrite(config['CREDENTIALS']['VPN_PASSWORD'])
    pyautogui.moveRel(0,60)
    pyautogui.click()
    
    vpn_connected = pyautogui.locateOnScreen('locations/vpn_connected.png')
    retry = 0
    while not vpn_connected:
        time.sleep(30)
        vpn_connected = pyautogui.locateOnScreen('locations/vpn_connected.png')
        retry += 1
        if retry == 6:
            raise ValueError("wait for connected more than 3 minutes, abort the job")
    print('SUCCESS : Connected to VPN')

if __name__ == "__main__":
    main()