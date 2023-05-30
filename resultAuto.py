import pyautogui
import pyperclip
import time

import pandas
import os

#goto firefox and click to make it active window
time.sleep(1)
pyautogui.moveTo(317,535)
pyautogui.click()



for i in range(57,83):
     #step 13, refresh page
    pyautogui.moveTo(41,657)
    pyautogui.click()
    pyautogui.hotkey('ctrl','r')
    time.sleep(8)

    #Step 0, click on process single station
    pyautogui.moveTo(56,141)
    pyautogui.click()
    time.sleep(1)

    #Step1, browse and put file name
    pyautogui.moveTo(172,255)
    pyautogui.click()
    filename=f"hist_ssp370_Station_{i}.txt"
    pyperclip.copy(filename)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(2)

    #Step 3, enter lati
    import openpyxl
    workbook = openpyxl.load_workbook('latLongData.xlsx')
    worksheet = workbook['Sheet1']

    cell_address = f'A{i+1}'
    cell_value = worksheet[cell_address].value
    variable = cell_value

    cell_address2 = f'B{i+1}'
    cell_value2 = worksheet[cell_address2].value
    variable2 = cell_value2


    lati=variable
    longi=variable2
    start=1951
    end=2014

    pyautogui.moveTo(200,358)
    pyautogui.click()
    pyautogui.click()
    pyautogui.click()
    pyperclip.copy(lati)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')

    #Step 4, enter longi
    pyautogui.moveTo(200,391)
    pyautogui.click()
    pyautogui.click()
    pyautogui.click()
    pyperclip.copy(longi)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')

    #Step 5, enter base start
    pyautogui.moveTo(200,435)
    pyautogui.click()
    pyautogui.click()
    pyautogui.click()
    pyperclip.copy(start)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')

    #Step 6, enter base end
    pyautogui.moveTo(200,464)
    pyautogui.click()
    pyautogui.click()
    pyautogui.click()
    pyperclip.copy(end)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(8)


    #Step 7,click next :: Load tab completed
    pyautogui.moveTo(459,521)
    pyautogui.click()
    time.sleep(1)

    #Step 8,click check data quality
    pyautogui.moveTo(327,232)
    pyautogui.click()
    time.sleep(50)

    #Step 9,click here
    time.sleep(3)
    pyautogui.moveTo(219,204)
    pyautogui.click()
    pyautogui.click()
    time.sleep(1) #wait 8 sec for qc data to download

    #Step 10,click next
    pyautogui.moveTo(450,531)
    pyautogui.click()
    time.sleep(1) 

    #Step 11,calculate indices
    time.sleep(1)
    pyautogui.moveTo(299,620)
    pyautogui.click()
    time.sleep(54)

    #Step 12,click here after indices
    time.sleep(1)
    pyautogui.moveTo(241,325)
    pyautogui.click()
    time.sleep(1)


    time.sleep(2)








