from pyautocad import Autocad
import os
import psutil
import time
import tkinter as tk
from tkinter import messagebox

# Global variables
running = True

# Exports data from pages 29, 30, 32 and converts to csv
def export(dir, jobNum):
    # Directory variables
    scriptDir = os.getcwd() + '\\AutoCAD Script\\Synergy Semi Auto Scripts'

    ############################
    # Export sheet 29 tables   #
    ############################
    file = dir + '\\' + jobNum + "_Sheet29.dwg"
    if running == True:
        os.startfile(file)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            # Show user info about how to save table 29a
            tk.messagebox.showinfo(title="Important", message="Please save file as 'JobNumber_Table29a.csv' and keep csv files together for importing to Access.")

            # Page 29 export script A
            while True:
                if running == True:
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-exportScript.scr \n')
                        break
                    except: pass
                else: return

            # Show user info about how to save table 29b
            tk.messagebox.showinfo(title="Important", message="Please save file as 'JobNumber_Table29b.csv' and keep csv files together for importing to Access.")

            # Page 29 export script B
            while True:
                if running == True:
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-exportScriptB.scr \n')
                        break
                    except: pass
                else: return
    else: return
    
    ###########################
    # Exoprt sheet 30 table   #
    ###########################
    file = dir + '\\' + jobNum + '_Sheet30.dwg'
    if running == True:
        os.startfile(file)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            # Show user info about how to save table 30
            tk.messagebox.showinfo(title="Important", message="Please save file as 'JobNumber_Table30.csv' and keep csv files together for importing to Access.")

            # Page 30 export script
            while True:
                if running == True:
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\30-exportScript.scr \n')
                        break
                    except: pass
                else: return
    else: return

    ###########################
    # Export sheet 32 table   #
    ###########################
    file = dir + '\\' + jobNum + '_Sheet32.dwg'
    if running == True:
        os.startfile(file)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            # Show user info about how to save table 32
            tk.messagebox.showinfo(title="Important", message="Please save file as 'JobNumber_Table32.csv' and keep csv files together for importing into Access.")

            # Page 32 export script
            while True:
                if running == True:
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-exportScript.scr \n')
                        break
                    except: pass
                else: return
    else: return

    # Show user that csv fiels were created and prompts them to move them into jobNum Electrical BOM folder
    tk.messagebox.showinfo(title="Table Data Files Created", message="CSV files for job " + jobNum + " were created where you saved them, please move them to '" + jobNum + " Electrical BOM' in Jobs directory.")
    