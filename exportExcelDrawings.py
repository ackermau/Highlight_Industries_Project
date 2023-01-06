from pyautocad import Autocad
from time import sleep
import os
import psutil
import keyboard
import threading
import time
import tkinter as tk
from tkinter import messagebox

# Script function
def scriptFunc(acad, scriptDir, page):
    if page == 29:
        timer = time.perf_counter()
        while True:
            endTimer = time.perf_counter()
            if (endTimer - timer) >= 60:
                return
            try:
                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-exportScript.scr \n')
                break
            except:
                pass
        timer = time.perf_counter()
        while True:
            endTimer = time.perf_counter()
            if (endTimer - timer) >= 60:
                return
            try:
                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-exportScriptB.scr \n')
                break
            except:
                pass
    elif page == 30:
        timer = time.perf_counter()
        while True:
            endTimer = time.perf_counter()
            if (endTimer - timer) >= 60:
                return
            try:
                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\30-exportScript.scr \n')
                break
            except:
                pass
    else:
        timer = time.perf_counter()
        while True:
            endTimer = time.perf_counter()
            if (endTimer - timer) >= 60:
                return
            try:
                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-exportScript.scr \n')
                break
            except:
                pass

# Keyboard function
def keyFunc(mainDir, jobNum, page, delay):
    if page == 29:
        sleep(20 + delay)
        keyboard.write(jobNum + '_Table29a.csv')
        sleep(1)
        keyboard.send('enter')
        sleep(delay)
        keyboard.write(jobNum + '_Table29b.csv')
        sleep(1)
        keyboard.send('enter')
    elif page == 30:
        sleep(delay)
        keyboard.write(jobNum + '_Table30.csv')
        sleep(1)
        keyboard.send('enter')
    else:
        sleep(delay)
        keyboard.write(jobNum + '_Table32.csv')
        sleep(1)
        keyboard.send('enter')

# Exports data from pages 29, 30, and 32 
def export(dir, jobNum, delay):
    scriptDir = os.getcwd() + '\\AutoCAD Script\\Synergy Semi Auto Scripts'
    mainDir = os.getcwd()

    # Export sheet 29 tables
    file = dir + "\\" + jobNum + "_Sheet29.dwg"
    page = 29
    os.startfile(file)
    if "acad.exe" in (i.name() for i in psutil.process_iter()):
        acad = Autocad()
        # creating threads for page 29
        sThread = threading.Thread(target=scriptFunc, args=(acad, scriptDir, page))
        kThread = threading.Thread(target=keyFunc, args=(mainDir, jobNum, page, delay))
        
        # starting script thread
        sThread.start()
        # starting keyboard thread
        kThread.start()

        # wait until script tread is complete
        sThread.join()
        # wait until keyboard thread is complete
        kThread.join()
        # keyboard.add_hotkey('backspace', keyFunc, args =(mainDir, page, table))
    # autocad not running error
    else:
        pass
    
    # Export sheet 30 tables
    file = dir + "\\" + jobNum + "_Sheet30.dwg"
    page = 30
    os.startfile(file)
    if "acad.exe" in (i.name() for i in psutil.process_iter()):
        acad = Autocad()
        # creating threads for page 30
        sThread = threading.Thread(target=scriptFunc, args=(acad, scriptDir, page))
        kThread = threading.Thread(target=keyFunc, args=(mainDir, jobNum, page, delay))

        # starting script thread
        sThread.start()
        # starting keyboard thread
        kThread.start()

        # wait until script thread is complete
        sThread.join()
        # wait until keyboard thread is complete
        kThread.join()
    # autocad not running error
    else:
        pass
    
    # Export sheet 32 tables
    file = dir + "\\" + jobNum + "_Sheet32.dwg"
    page = 32
    os.startfile(file)
    if "acad.exe" in (i.name() for i in psutil.process_iter()):
        acad = Autocad()
        # creating threads for page 32
        sThread = threading.Thread(target=scriptFunc, args=(acad, scriptDir, page))
        kThread = threading.Thread(target=keyFunc, args=(mainDir, jobNum, page, delay))

        # starting script thread
        sThread.start()
        # starting keyboard thread
        kThread.start()

        # wait until script thread is complete
        sThread.join()
        # wait until keyboard thread is complete
        kThread.join()
    # autocad not running error
    else:
        pass

    # message sent to user to inform them csv files were created and prompts them to move them into the job directory
    tk.messagebox.showwarning(title="Table Data Files Created", message="CSV files of job " + jobNum + " were created where your AutoCAD saved them, please move them into Jobs database")