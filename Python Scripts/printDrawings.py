from pyautocad import Autocad
import os
import psutil
import time
import tkinter as tk
from tkinter import messagebox

# Global variables
running = True

# Prints all pages desired from a specific job
def printAllDrawings(jobNum, dir):
    scriptDir = os.getcwd() + '\\AutoCAD Script\\Synergy Semi Auto Scripts'
    # getting only dwg files
    fileList = os.listdir(dir)
    for file in fileList:
        if not file.split(".")[1] == "dwg":
            fileList.remove(file)
    # geting number of files in directory to print
    numFiles = len(fileList)
    # send commands to print all of the sheets
    for x in range(0, numFiles):
        # Sheets 00 - 09
        if running == True:
            if x >= 0 and x <= 9:
                file = dir + "\\" + jobNum + "_Sheet0" + str(x) + ".dwg"
                if running == True:
                    os.startfile(file)
                    if "acad.exe" in (i.name() for i in psutil.process_iter()):
                        acad = Autocad()
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\newPrintScript.scr \n')
                                break
                            except: pass
                    # autocad not running error
                    else: pass
                else: return
            # Rest of Sheets
            else:
                file = dir + "\\" + jobNum + "_Sheet" + str(x) + ".dwg"
                if running == True:
                    os.startfile(file)
                    if "acad.exe" in (i.name() for i in psutil.process_iter()):
                        acad = Autocad()
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\newPrintScript.scr \n')
                                break
                            except: pass
                    # autocad not running error
                    else: pass
                else: return
        else: return
    
    # End scripts
    file = dir + "\\" + jobNum + "_Sheet00.dwg"
    if running == True:
        os.startfile(file)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
            timer = time.perf_counter()
            while True:
                endTimer = time.perf_counter()
                if (endTimer - timer) >= 60: return
                try:
                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\endScrtip.scr \n')
                    break
                except: pass
        # autocad not running error
        else: pass
    else: return

    # sends user message telling them all pages of jobNum were printed
    tk.messagebox.showinfo(title="Printed Job " + jobNum, message="All pages of job " + jobNum + " was printed")


# prints selected pages desired from a specific job
def printSelDrawings(jobNum, dir, pages):
    scriptDir = os.getcwd() + '\\AutoCAD Script\\Synergy Semi Auto Scripts'
    pagesWRanges = pages.split(",")
    for p in pagesWRanges:
        if running == True:
            # multiple sheet print
            if "-" in p:
                miniRange = p.split("-")
                for x in range(int(miniRange[0]), int(miniRange[1]) + 1):
                    if x >= 0 and x <= 9:
                        file = dir + "\\" + jobNum + "_Sheet0" + str(x) + ".dwg"
                        if running == True:
                            os.startfile(file)
                            if "acad.exe" in (i.name() for i in psutil.process_iter()):
                                acad = Autocad()
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\newPrintScript.scr \n')
                                        break
                                    except: pass
                            # autocad not running error
                            else: pass
                        else: return
                    else:
                        file = dir + "\\" + jobNum + "_Sheet" + str(x) + ".dwg"
                        if running == True:
                            os.startfile(file)
                            if "acad.exe" in (i.name() for i in psutil.process_iter()):
                                acad = Autocad()
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\newPrintScript.scr \n')
                                        break
                                    except: pass
                            # autocad not running error
                            else: pass
                        else: return
            # single sheet print                
            else:
                if int(p) >= 0 and int(p) <= 9:
                        file = dir + "\\" + jobNum + "_Sheet0" + p + ".dwg"
                        if running == True:
                            os.startfile(file)
                            if "acad.exe" in (i.name() for i in psutil.process_iter()):
                                acad = Autocad()
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\newPrintScript.scr \n')
                                        break
                                    except: pass
                            # autocad not running error
                            else: pass
                        else: return
                else:
                    file = dir + "\\" + jobNum + "_Sheet" + p + ".dwg"
                    if running == True:
                        os.startfile(file)
                        if "acad.exe" in (i.name() for i in psutil.process_iter()):
                            acad = Autocad()
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\newPrintScript.scr \n')
                                    break
                                except: pass
                        # autocad not running error
                        else: pass
                    else: return
        else: return

    # End scripts
    file = dir + "\\" + jobNum + "_Sheet00.dwg"
    if running == True:
        os.startfile(file)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
            timer = time.perf_counter()
            while True:
                endTimer = time.perf_counter()
                if (endTimer - timer) >= 60: return
                try:
                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\endScript.scr \n')
                    break
                except: pass
        # autocad not running error
        else: pass
    else: return

    # sends user message telling them the selected pages were printed from job jobNum
    tk.messagebox.showinfo(title="Printed Job " + jobNum, message="Pages " + pages + " of Job " + jobNum + " was printed")
