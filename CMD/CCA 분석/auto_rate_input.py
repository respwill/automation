# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 15:27:30 2019

@author: parkhi
"""

# need to open SAP 

from pywinauto.application import Application
import pyautogui
import time
import datetime
import pandas as pd
import sqlite3
import win32clipboard
from tqdm import tqdm
import ctypes
#import shutil

class cca_rate_input():
    def __init__(self, target_q_period, data_file, act_capa):
        
        # determine if keep activity and capacity 
        self.act_capa = act_capa
        
        # set period by targt quarter
        target_q = target_q_period.lower()
        
        now = datetime.datetime.now()
        self.year = now.strftime('%Y')
        
        if target_q == 'q1':
            self.start_month = '1'
            self.end_month = '3'
#            self.year = str(int(self.year) + 1)
            self.year = str(int(self.year))
        elif target_q == 'q2':
            self.start_month = '4'
            self.end_month = '6'
        elif target_q == 'q3':
            self.start_month = '7'
            self.end_month = '9'
        else:
            self.start_month = '10'
            self.end_month = '12'
            
        self.data = pd.read_excel(data_file)
        
        #decimal points of rate should be shorter than 3
        self.data = self.data.round(2)
        
        self.mc = self.data[(self.data['type']=='mach_s')]
        self.lb = self.data[(self.data['type']=='lab_s')]
        
        

    def SAP_log_in(self, server):
        conn = sqlite3.connect("D:\BACK UP\Password DB\password.db")
        cur = conn.cursor()
        if server == 'prd':
            cur.execute("select pswd from USER_INFO where system='PRD'")
            self.password = cur.fetchall()[0][0]
        else:
            cur.execute("select pswd from USER_INFO where system='TQA'")
            self.password = cur.fetchall()[0][0]
            
            
        app = Application()
        app.start(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
    
        time.sleep(10)
        
        # PRD module is at first row so press enter is fine.
        if server == 'prd':
            pyautogui.press('enter')
        else:
            for n in range(0,3):
                pyautogui.press('down')
            pyautogui.press('enter')
    
        time.sleep(5)
    
        # ID / PW 입력
        pyautogui.typewrite('parkhi')
        pyautogui.press('tab')
        pyautogui.typewrite(self.password)
        pyautogui.press('enter')
        # SAP code 입력
        
        time.sleep(5)
    
    def run_kp26_cca(self):
        pyautogui.typewrite('KP26')
        pyautogui.press('enter')
        
        time.sleep(2)
        
        # input control area
        pyautogui.typewrite('1000')
        pyautogui.press('enter')
        
        time.sleep(2)
        
        # input version as 0
        pyautogui.typewrite('0')
        pyautogui.press('tab')
        
        # input period information
        pyautogui.typewrite(self.start_month)
        pyautogui.press('tab')
        pyautogui.typewrite(self.end_month)
        pyautogui.press('tab')
        pyautogui.typewrite(self.year)
        pyautogui.press('tab')
        
        #input rate for machine and labor
    
    def input_kp26_cca(self, target):
#        print(self.mc)
        if target == 'mc':
            df = self.mc
            target_type = 'mach_s'
        else:
            df = self.lb
            target_type = 'lab_s'
        
        
        for n in tqdm(range(0, len(df))):
            row = df.iloc[n]
#            print(row['Cost Center'], row['Fixed'], row['Variable'])
            fix = row['Fixed']
            var = row['Variable']
            cc = row['Cost Center']
#            print(cc)

            # delete cc which already input.
            pyautogui.press('del')
            pyautogui.typewrite(cc)
            for n in range(0,3):
                pyautogui.press('tab')
            pyautogui.press('del')
            pyautogui.typewrite(target_type)
            pyautogui.hotkey('f6')
            time.sleep(2)
            
            # page changed to input area
            # copy the plan activity rate
            
            # get value from clipboard to see if they are blank.
            for n in range(0,3):
                # in case plan activity quantity or capacity alraedy exists, remain it otherwise input 1
                
                for n in ['Plan activity', 'Capacity']:
                    if self.act_capa == True:
                        pyautogui.hotkey('ctrl', 'c')
                        win32clipboard.OpenClipboard()
                        
                        try:
                            plan_act = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        except (TypeError, win32clipboard.error):
                            try:
                                plan_act = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
    #                            text = py3compat.cast_unicode(text, py3compat.DEFAULT_ENCODING)
                            except (TypeError, win32clipboard.error):
                                plan_act = '0'
    
                        win32clipboard.EmptyClipboard()
                        win32clipboard.CloseClipboard()
                        
    #                    print("{} of {}:".format(n, cc), plan_act)
    #                    print(str(plan_act))
                        if (str(plan_act) == 'return 0') | (str(plan_act) == '0'):
                            pyautogui.typewrite('1')
                            pyautogui.press('tab')
                        else:
                            pyautogui.press('tab')
                    
                    else:
                        pyautogui.typewrite('1')
                        pyautogui.press('tab')
                    

                for d in [fix, var, '10000']:
                    if str(d) == 'nan':
                        pyautogui.press('tab')
                    else:
                        pyautogui.press('del')
                        pyautogui.typewrite(str(d))
                        pyautogui.press('tab')
                    
            
                pyautogui.press('tab')
                

            # save the data that input     
            
            print('Input rate for {} completed: fix:{}, var:{}'.format(cc, fix, var))
            pyautogui.hotkey('ctrl', 's')
            
            time.sleep(5)
            # go to Cost Center tap
            for n in range(0,4):
                pyautogui.press('tab')
                

#********************************************************************
cca = cca_rate_input('q1', 'Manual Input Rate_1Q20_python_tqa.xlsx', False)
cca.SAP_log_in('tqa')
cca.run_kp26_cca()
cca.input_kp26_cca('mc')
cca.input_kp26_cca('lb')


