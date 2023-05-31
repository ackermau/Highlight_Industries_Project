from pyautocad import Autocad
import os
import psutil
import time
import tkinter as tk
from tkinter import messagebox

# Global variables
running = True

# Prints all pages from a specific job
def printAllDrawings(jobNum, dir):
    scriptDir = os.getcwd() + '\\AutoCAD Script\\Synergy Automatic Scripts'
    # getting dwg file
    if running == True:
        file = dir + "\\" + jobNum + " Schematics.dwg"
        os.startfile(file)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
            while True:
                try:
                    acad.doc.SendCommand('FILEDIA 0 \n')
                    break
                except: pass
            # Pages 00 - 09
            for x in range(10):
                layout = "PAGE 0" + str(x)
                # setting layout for pages 0 - 9 for newPrintScript
                with open(scriptDir + '\\newPrintScript.scr', 'r') as tempFile:
                    script = tempFile.read()
                    script = script.replace('<LAYOUT>', layout)
                with open(scriptDir + '\\page0' + str(x) + 'PrintScript.scr', 'w') as tempFile:
                    tempFile.write(script)
                if running == True:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\page0' + str(x) + 'PrintScript.scr \n')
                            break
                        except: pass
                else: return
            for x in range(10, 30):
                layout = "PAGE " + str(x)
                # setting layout for pages 10 - 32 for newPrintScript
                with open(scriptDir + '\\newPrintScript.scr', 'r') as tempFile:
                    script = tempFile.read()
                    script = script.replace('<LAYOUT>', layout)
                with open(scriptDir + '\\page' + str(x) + 'PrintScript.scr', 'w') as tempFile:
                    tempFile.write(script)
                if running == True:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\page' + str(x) + 'PrintScript.scr \n')
                            break
                        except: pass
                else: return
            
            # End Scripts
            timer = time.perf_counter()
            while True:
                endTimer = time.perf_counter()
                if (endTimer - timer) >= 60: return
                try:
                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\endScript.scr \n')
                    break
                except: pass
        else: return
    else: return

    # sends user message telling them all pages of jobNum were printed
    tk.messagebox.showinfo(title="Printed Job " + jobNum, message="All pages of job " + jobNum + " was printed")

# prints selected pages desired for a specific job
def printSelDrawings(jobNum, dir, pages):
    scriptDir = os.getcwd() + '\\AutoCAD Script\\Synergy Automatic Scripts'
    pagesWRanges = pages.split(",")
    file = dir + "\\" + jobNum + " Schematics.dwg"
    os.startfile(file)
    if "acad.exe" in (i.name() for i in psutil.process_iter()):
        acad = Autocad()
        for p in pagesWRanges:
            if running == True:
                # multiple sheet print
                if "-" in p:
                    miniRange = p.split("-")
                    for x in range(int(miniRange[0]), int(miniRange[1]) + 1):
                        if x >= 0 and x <= 9:
                            if running == True:
                                layout = "PAGE 0" + str(x)
                                with open(scriptDir + '\\newPrintScript.scr', 'r') as tempFile:
                                    script = tempFile.read()
                                    script = script.replace('<LAYOUT>', layout)
                                with open(scriptDir + '\\page0' + str(x) + 'PrintScript.scr', 'w') as tempFile:
                                    tempFile.write(script)
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\page0' + str(x) + 'PrintScript.scr \n')
                                        break
                                    except: pass
                            else: return
                        else:
                            if running == True:
                                layout = "PAGE " + str(x)
                                with open(scriptDir + '\\newPrintScript.scr', 'r') as tempFile:
                                    script = tempFile.read()
                                    script = script.replace('<LAYOUT>', layout)
                                with open(scriptDir + '\\page' + str(x) + 'PrintScript.scr', 'w') as tempFile:
                                    tempFile.write(script)
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\page' + str(x) + 'PrintScript.scr \n')
                                        break
                                    except: pass 
                            else: return
                # single sheet print
                else:
                    if int(p) >= 0 and int(p) <= 9:
                        if running == True:
                            layout = "PAGE 0" + str(p)
                            with open(scriptDir + '\\newPrintScript.scr', 'r') as tempFile:
                                script = tempFile.read()
                                script = script.replace('<LAYOUT>', layout)
                            with open(scriptDir + '\\page0' + str(p) + 'PrintScript.scr', 'w') as tempFile:
                                tempFile.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\page0' + str(p) + 'PrintScript.scr \n')
                                    break
                                except: pass
                        else: return
                    else:
                        if running == True:
                            layout = "PAGE " + str(p)
                            with open(scriptDir + '\\newPrintScript.scr', 'r') as tempFile:
                                script = tempFile.read()
                                script = script.replace('<LAYOUT>', layout)
                            with open(scriptDir + '\\page' + str(p) + 'PrintScript.scr', 'w') as tempFile:
                                tempFile.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\page' + str(p) + 'PrintScript.scr \n')
                                    break
                                except: pass
                        else: return
            else: return
        
        # End Scripts
        timer = time.perf_counter()
        while True:
            endTimer = time.perf_counter()
            if (endTimer - timer) >= 60: return
            try:
                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\endScript.scr \n')
                break
            except: pass

        # sends user message telling them the selected pages were printed from job jobNum
        tk.messagebox.showinfo(title="Printed Job " + jobNum, message="Pages " + pages + " of Job " + jobNum + " was printed")