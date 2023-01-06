import tkinter as tk
import os
import shutil
import drawingsEdit
import printDrawings
import exportExcelDrawings
import exExcelMan
import directoryFinder
import exportAccessDrawings
import drawingsEditLayers
import re
import threading
from tkinter import *
from tkinter import messagebox

########################################
# Reset function to make gui default   #
########################################
def resetCallBack(macVar, synVar, profileVar, scaleVar, coldVar, splitVar, autoVar, ulVar, doorVar, harVar,
        custEntry, distrEntry, projNumEntry, manYearEntry, enginEntry, dateEntry,
        phaseEntry, mainLineVEntry, controlVEntry, totMotorEntry, fullLoadEntry):
    # resetting gui to default values
    # Customer Entry
    custEntry.delete(0, tk.END)

    # Distributor Entry
    distrEntry.delete(0, tk.END)

    # Project Entry
    projNumEntry.delete(0, tk.END)

    # Manufacturing Year Entry
    manYearEntry.delete(0, tk.END)

    # Engineer Entry
    enginEntry.delete(0, tk.END)

    # Date Entry
    dateEntry.delete(0, tk.END)

    # Variables
    macVar.set(0)
    synVar.set(0)
    profileVar.set(0) 
    scaleVar.set(0)
    coldVar.set(0)
    splitVar.set(0)
    autoVar.set(0)
    ulVar.set(0)
    doorVar.set(0)
    harVar.set(0)

    # Phase Entry
    phaseEntry.delete(0, tk.END)
    phaseEntry.insert(0, "60 HERTZ")

    # Main Line Entry
    mainLineVEntry.delete(0, tk.END)
    mainLineVEntry.insert(0, "120 VAC, 1 PHASE")

    # Control Voltage Entry
    controlVEntry.delete(0, tk.END)
    controlVEntry.insert(0, "120 VAC, 24VDC")

    # Total Motor Entry
    totMotorEntry.delete(0, tk.END)
    totMotorEntry.insert(0, "2 HP")

    # Full Load Entry
    fullLoadEntry.delete(0, tk.END)
    fullLoadEntry.insert(0, "15 AMP")

    # Setting focus
    custEntry.focus()

######################################################
# Threading function for doneCallBack to help stop   #
######################################################
def doneThreadFunc(destFolder, custText, distrText, projNumText, manYearText, 
        phaseText, mainVText, controlVText, totMotorText, fullLoadText,
        enginText, dateText,
        synVar, profileVar, scaleVar, coldVar, splitVar, autoVar, ulVar, doorVar, harVar):
    # calling drawings edit method to update sheets to user input
    drawingsEditLayers.editLayers(destFolder, custText, distrText, projNumText, manYearText, phaseText, mainVText, controlVText, totMotorText, fullLoadText, enginText, dateText, synVar, profileVar, scaleVar, coldVar, splitVar, autoVar, ulVar, doorVar, harVar)

##################################################################################
# method that is called when user has entered all fields and clicks done button  #
##################################################################################
def doneCallBack(macVar, 
        custEntry, distrEntry, projNumEntry, manYearEntry,
        phaseEntry, mainLineVEntry, controlVEntry, totMotorEntry, fullLoadEntry, 
        enginEntry, dateEntry,
        synVar, profileVar, scaleVar, coldVar, splitVar, autoVar, ulVar, doorVar, harVar,
        numRan):
    if macVar.get() == 1:
        tk.messagebox.showerror(title="Work in Progress", message="Still working on Freedom machine")
        return

    # variables and checks
    # Customer variable and check
    custText = custEntry.get()
    if len(custText) > 25:
        tk.messagebox.showerror(title="Invalid Entry", message="Customer name exceeds given limit of 25 characters")
        return

    # Distributor variable and check
    distrText = distrEntry.get()
    if len(distrText) > 25:
        tk.messagebox.showerror(title="Invalid Entry", message="Distributor name exceeds given limit of 25 characters")
        return

    # Project number variable and check
    projNumText = projNumEntry.get()
    if len(projNumText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Project number exceeds given limit of 15 characters")
        return

    # Manufacturer year variable and check
    manYearText = manYearEntry.get()
    if len(manYearText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Manufacturing year exceeds given limit of 15 characters")
        return

    # Phase variable and check
    phaseText = phaseEntry.get()
    if len(phaseText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Phase entry exceeds given limit of 15 characters")
        return

    # Main line voltage variable and check
    mainLineVText = mainLineVEntry.get()
    if len(mainLineVText) > 18:
        tk.messagebox.showerror(title="Invalid Entry", message="Main line voltage exceeds given limit of 15 characters")
        return

    # Control Voltage variable and check
    controlVText = controlVEntry.get()
    if len(controlVText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Control voltage exceeds given limit of 15 characters")
        return

    # Total motor variable and check
    totMotorText = totMotorEntry.get()
    if len(totMotorText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Total Motor entry exceeds given limit of 15 characters")
        return

    # Full load variable and check
    fullLoadText = fullLoadEntry.get()
    if len(fullLoadText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Full load entry exceeds given limit of 15 characters")
        return 

    # Engineer initials variable and check
    enginText = enginEntry.get()
    if len(enginText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Engineer initials exceeds the given limit of 15 characters")
        return

    # Date variable and check - format ##/##/####
    dateText = dateEntry.get()
    if len(dateText) > 15:
        tk.messagebox.showerror(title="Invalid Entry", message="Date exceeds the given limit of 15 characters")
        return
    reDateText = re.search("(?:0[1-9]|1[012])/(?:0[1-9]|[12][0-9]|3[01])/(?:19\d{2}|20[0-9][0-9])", dateText)
    if not reDateText:
        tk.messagebox.showerror(title="Invalid Entry", message="Date format is MM/DD/YYYY")
        return

    # Folder variables
    sourceFolder = "G:\\~~EE Directory~~\\01-AutoCAD Electrical\\07-Standard Drawings Templates\\Synergy Semi Auto"
    destinationFolder = os.getcwd() + "\\" + projNumText + " Drawings"

    # new folder destination
    try: 
        os.mkdir(destinationFolder)
    except:
        res = tk.messagebox.askyesno(title="Job Already Exists", message="Job already exists in directory, replace?")
        if res == True:
            shutil.rmtree(destinationFolder)
            os.mkdir(destinationFolder)
        else: return

    # Copies original schematic to new folder
    for fileName in os.listdir(sourceFolder):
        source = sourceFolder + "\\" + fileName
        destination = destinationFolder + "\\" + fileName
        if os.path.isfile(source):
            shutil.copy(source, destination)

    # Renames new folders files to match user inputs
    for fileName in os.listdir(destinationFolder):
        files = fileName.split(".")
        if files[1] == "aepx":
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + projNumText + " Drawings.aepx")
        elif files[1] == "wdp":
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + projNumText + " Drawings.wdp")
        elif files[1] == "dwg":
            noNumFile = fileName.split("_")
            noNumFileNoEx = noNumFile[1].split(".")
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + projNumText + "_" + noNumFileNoEx[0] + ".dwg")
        elif files[1] == "bak":
            noNumFile = fileName.split("_")
            noNumFileNoEx = noNumFile[1].split(".")
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + projNumText + "_" + noNumFileNoEx[0] + ".bak")
        else: pass

    # setting running to true for drawings edit script
    drawingsEditLayers.running = True

    # creation of edit thread to integrate stop button
    global editThread
    editThread = threading.Thread(target=doneThreadFunc, args=(destinationFolder, custText, distrText, projNumText, manYearText,
                                                            phaseText, mainLineVText, controlVText, totMotorText, fullLoadText, 
                                                            enginText, dateText, 
                                                            synVar.get(), profileVar.get(), scaleVar.get(), coldVar.get(), splitVar.get(), autoVar.get(), ulVar.get(), doorVar.get(), harVar.get()))
    
    # check to see if program has been stopped or not
    if numRan == 0:
        editThread.start()
        numRan = 1
    else:
        editThread = threading.Thread(target=doneThreadFunc, args=(destinationFolder, custText, distrText, projNumText, manYearText, 
                                                            phaseText, mainLineVText, controlVText, totMotorText, fullLoadText, 
                                                            enginText, dateText, 
                                                            synVar.get(), profileVar.get(), scaleVar.get(), coldVar.get(), splitVar.get(), autoVar.get(), ulVar.get(), doorVar.get(), harVar.get()))
        editThread.start()

#################################################################
# Method called when user presses stop button on main tab       #
#################################################################
def stopCallBack(projNumEntry):
    drawingsEditLayers.running = False
    tk.messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopped editing " + projNumEntry.get() + " Drawings.")
    editThread.join(3)

########################################################
# Threading function for exportCallBack to help stop   #
########################################################
def exportThreadFunc(jobNum, dir, mDir, exportJobVar):
    # Case 1: To excel with automatic finder
    if exportJobVar.get() == 0:
        exExcelMan.export(dir, jobNum)

    # Case 2: To excel with manual directory
    else:
        exExcelMan.export(mDir, jobNum)

#########################################################################
# Method called when user fills out export tab and presses export       #
#########################################################################
def exportCallBack(exJobNumEntry, exDirEntry, exportJobVar, numRan):
    # variables and checks
    jobNum = exJobNumEntry.get()
    dir = directoryFinder.findDir(jobNum)
    mDir = exDirEntry.get()

    #
    # general error proof checks
    #
    # Job number cannot be empty
    if jobNum == "":
        tk.messagebox.showerror(title="Job Number Empty", message="Job number field empty, please enter job number")
        return
    
    # Check to confirm manual directory is not empty when using manual directory
    if mDir == "" and exJobNumEntry.get() == 1:
        tk.messagebox.showerror(title="Manual Directory Empty", message="Manual directory field empty, please enter manual directory")
        return

    # Check to make sure automatic directory exists
    testFile = dir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(testFile) == False and exJobNumEntry.get() == 0:
        tk.messagebox.showerror(title="Invalid Directory", message="Automatic finder could not find job number directory, please enter manual directory")
        return

    # Check to make sure manual directory exists
    mTestFile = mDir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(mTestFile) == False and exJobNumEntry.get() == 1:
        tk.messagebox.showerror(title="Invalid Directory", message="Manual directory is invalid, please try again")
        return

    # Setting running to true for export drawings
    exExcelMan.running = True

    # Creation of export thread to integrate stop button
    global exportThread
    exportThread = threading.Thread(target=exportThreadFunc, args=(jobNum, dir, mDir, exportJobVar))

    # Check to see if the program has been stopped or not
    if numRan == 0:
        exportThread.start()
        numRan = 1
    else:
        exportThread = threading.Thread(target=exportThreadFunc, args=(jobNum, dir, mDir, exportJobVar))
        exportThread.start()

###############################################################
# Method called when user presses stop button in export tab   #
###############################################################
def exStopCallBack(exJobNumEntry):
    exExcelMan.running = False
    tk.messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopped exporting " + exJobNumEntry.get() + " Drawings.")

##########################################################################
# Method called when user fills out the access tab and presses access    #
##########################################################################
def accessCallBack(accessJobEntry, accessDirEntry, accessJobVar):
    # variables and checks
    jobNum = accessJobEntry.get()
    dir = directoryFinder.findDir(jobNum)
    mDir = accessDirEntry.get()

    #
    # general error proof checks
    #
    # Job number cannot be empty
    if jobNum == "":
        tk.messagebox.showerror(title="Job Number Empty", message="Job number field empty, please enter job number")
        return

    # Check to confirm manual directory is not empty when using manual directory
    if mDir == "" and accessJobVar.get() == 1:
        tk.messagebox.showerror(title="Manual Directory Empty", message="Maunal directory field empty, please enter manual directory")
        return
    
    # Check to make sure automatic directory exists
    testFile = dir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(testFile) == False and accessJobVar.get() == 0:
        tk.messagebox.showerror(title="Invalid Directory", message="Automtic finder could not find job number directory, please enter maunal directory")
        return
    
    # Check to make sure maunal directory exists
    if os.path.exists(mDir) == False and accessJobVar.get() == 1:
        tk.messagebox.showerror(title="Invalid Directory", message="Manual directory is invalid, please try again")
        return
    
    # Case 1: To access with automatic finder
    if accessJobVar.get() == 0:
        exportAccessDrawings.toAccess(dir, jobNum)

    # Case 2: To access with maunal directory
    else:
        exportAccessDrawings.toAccess(mDir, jobNum)

#######################################################
# Threading function for printCallBack to help stop   #
#######################################################
def printThreadFunc(jobNum, jobVar, dir, mDir, pages, paVar):
    # Calling print drawings method to user specified pages
    # case 1: automatic finder with print all pages
    if jobVar.get() == 0 and paVar.get() == 1:
        printDrawings.printAllDrawings(jobNum, dir)

    # case 2: manual directory with print all pages
    elif jobVar.get() == 1 and paVar.get() == 1:
        printDrawings.printAllDrawings(jobNum, mDir)

    # case 3: automatic finder with selected pages
    elif jobVar.get() == 0 and paVar.get() == 0:
        printDrawings.printSelDrawings(jobNum, dir, pages)

    # case 4: manual directory with selected pages
    elif jobVar.get() == 1 and paVar.get() == 0:
        printDrawings.printSelDrawings(jobNum, mDir, pages)
    
    # error
    else:
        tk.messagebox.showerror(title="Error", message="This error should not happen")
        return

########################################################################
# Method called when user fills out the print tab and presses print    #
########################################################################
def printCallBack(dirEntry, jobEntry, pagesEntry, printerVar, jobVar, paVar, numRan):
    # variables and checks
    mDir = dirEntry.get()
    jobNum = jobEntry.get()
    dir = directoryFinder.findDir(jobNum)
    pages = pagesEntry.get()
    pages.replace(' ', '')
    if printerVar.get() == "":
        tk.messagebox.showerror(title="No Printer Selected", message="Please select a printer")
        return
    else:
        printer = printerVar.get()

    #
    # general error proof checks
    #
    # checks for blank text boxes
    # jobNum can never be empty
    if jobNum == "":
        tk.messagebox.showerror(title="Job Number Empty", message="Job number field empty, please enter job number")
        return
    
    # check to confirm manual directory is not empty when using manual search
    if mDir == "" and jobVar.get() == 1:
        tk.messagebox.showerror(title="Manual Directory Empty", message="Manual directory field empty, please enter manual directory")
        return
    
    # check to confirm pages is not empty when printing selected pages
    if pages == "" and paVar.get() == 0:
        tk.messagebox.showerror(title="Pages Empty", message="Pages field empty, please enter desired pages to print")
        return

    # check to confirm pages is only integers
    tempPages = pages
    tempPages = tempPages.replace(',', '')
    tempPages = tempPages.replace('-', '')
    if tempPages.isnumeric() == False and paVar.get() == 0:
        tk.messagebox.showerror(title="Pages Not Numbers", message="Pages field contains invalid characters, please try again")
        return

    # check to confirm jobNum is valid
    jobCheck = int(jobNum) // 10
    if dir == "Maybe":
        if jobCheck.isnumeric() == False:
            tk.messagebox.showerror(title="Job Number Not Numbers", message="Job number field contains invalid characters, please try again")
            return
    
    # check to confirm job number directory exists
    testFile = dir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(testFile) == False and jobVar.get() == 0:
        tk.messagebox.showerror(title="Invalid Directory", message="Automatic finder could not find job number directory, please enter manual directory")
        return

    # check to confirm manual directory exists
    mTestFile = mDir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(mTestFile) == False and jobVar.get() == 1:
        tk.messagebox.showerror(title="Invalid Directory", message="Manual directory is invalid, please try again")
        return

    # check to confirm printer is not empty
    if printer == "":
        tk.messagebox.showerror(title="Printer Empty", message="Printer field is empty, please enter printer name")
        return

    if '\\\\server\\' in printer:
        printer = '\\\\\\\\server\\\\EE - Sharp MX-3501N'
    # setting printer in printing script
    with open(os.getcwd() + '\\AutoCAD Script\\Synergy Semi Auto Scripts\\printScript.scr', 'r') as file:
        script = file.read()
        script = script.replace('<PRINTER>', printer)
    with open(os.getcwd() + '\\AutoCAD Script\\Synergy Semi Auto Scripts\\newPrintScript.scr', 'w') as file:
        file.write(script)

    # Setting running to true for print drawings
    printDrawings.running = True

    # Creation of print thread to integrate stop button
    global printThread
    printThread = threading.Thread(target=printThreadFunc, args=(jobNum, jobVar, dir, mDir, pages, paVar))

    # Check to see if the program has been stopped or not
    if numRan == 0:
        printThread.start()
        numRan = 1
    else:
        printThread = threading.Thread(target=printThreadFunc, args=(jobNum, jobVar, dir, mDir, pages, paVar))
        printThread.start()

#############################################################
# Method called when use presses stop button on print tab   #
#############################################################
def printStopCallBack(jobEntry):
    printDrawings.running = False
    tk.messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopped printing " + jobEntry.get() + " Drawings.")
    printThread.join(3)
