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
import autoDrawingsEditLayers
import re
import threading
from tkinter import *
from tkinter import messagebox

#################################################
# On closing function when user closes window   #
#################################################
def onClosing(window):
    if messagebox.askokcancel(title="Quit", message="Do you want to quit?"):
        drawingsEditLayers.running = False
        autoDrawingsEditLayers.running = False
        window.destroy()

########################################
# Reset function to make gui default   #
########################################
def resetCallBack(semiMachine):
    # resetting gui to default values
    # Customer Entry
    semiMachine.custEntry.delete(0, tk.END)

    # Distributor Entry
    semiMachine.distrEntry.delete(0, tk.END)

    # Project Entry
    semiMachine.projNumEntry.delete(0, tk.END)

    # Manufacturing Year Entry
    semiMachine.manYearEntry.delete(0, tk.END)

    # Engineer Entry
    semiMachine.enginEntry.delete(0, tk.END)

    # Date Entry
    semiMachine.dateEntry.delete(0, tk.END)

    # Variables
    semiMachine.synVar.set(0)
    semiMachine.profileVar.set(0) 
    semiMachine.scaleVar.set(0)
    semiMachine.coldVar.set(0)
    semiMachine.AMP20V.set(0)
    semiMachine.splitVar.set(0)
    semiMachine.autoVar.set(0)
    semiMachine.ulVar.set(0)
    semiMachine.doorVar.set(0)
    semiMachine.harVar.set(0)

    # Phase Entry
    semiMachine.phaseEntry.delete(0, tk.END)
    semiMachine.phaseEntry.insert(0, "60 HERTZ")

    # Main Line Entry
    semiMachine.mainLineVEntry.delete(0, tk.END)
    semiMachine.mainLineVEntry.insert(0, "120 VAC, 1 PHASE")

    # Control Voltage Entry
    semiMachine.controlVEntry.delete(0, tk.END)
    semiMachine.controlVEntry.insert(0, "120 VAC, 24VDC")

    # Total Motor Entry
    semiMachine.totMotorEntry.delete(0, tk.END)
    semiMachine.totMotorEntry.insert(0, "2 HP")

    # Full Load Entry
    semiMachine.fullLoadEntry.delete(0, tk.END)
    semiMachine.fullLoadEntry.insert(0, "15 AMP")

    # Setting focus
    semiMachine.custEntry.focus()

######################################################
# Threading function for doneCallBack to help stop   #
######################################################
def doneThreadFunc(destFolder, semiMachine):
    # calling drawings edit method to update sheets to user input
    drawingsEditLayers.editLayers(destFolder, semiMachine.custEntry.get(), semiMachine.distrEntry.get(), semiMachine.projNumEntry.get(), semiMachine.manYearEntry.get(), semiMachine.phaseEntry.get(), 
                        semiMachine.mainLineVEntry.get(), semiMachine.controlVEntry.get(), semiMachine.totMotorEntry.get(), semiMachine.fullLoadEntry.get(), semiMachine.enginEntry.get(), semiMachine.dateEntry.get(), 
                        semiMachine.synVar.get(), semiMachine.profileVar.get(), semiMachine.scaleVar.get(), semiMachine.coldVar.get(), semiMachine.AMP20V.get(), semiMachine.splitVar.get(), semiMachine.autoVar.get(), 
                        semiMachine.ulVar.get(), semiMachine.doorVar.get(), semiMachine.harVar.get())

####################################################################################################
# method that is called when user has entered all fields and clicks done button for semi machines  #
####################################################################################################
def doneCallBack(semiMachine):
    # variables and checks
    # Customer check
    if len(semiMachine.custEntry.get()) > 25:
        messagebox.showerror(title="Invalid Entry", message="Customer name exceeds given limit of 25 characters")
        return

    # Distributor check
    if len(semiMachine.distrEntry.get()) > 25:
        messagebox.showerror(title="Invalid Entry", message="Distributor name exceeds given limit of 25 characters")
        return

    # Project number check
    if len(semiMachine.projNumEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Project number exceeds given limit of 15 characters")
        return

    # Manufacturer year check
    if len(semiMachine.manYearEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Manufacturing year exceeds given limit of 15 characters")
        return

    # Phase check
    if len(semiMachine.phaseEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Phase entry exceeds given limit of 15 characters")
        return

    # Main line voltage check
    if len(semiMachine.mainLineVEntry.get()) > 18:
        messagebox.showerror(title="Invalid Entry", message="Main line voltage exceeds given limit of 15 characters")
        return

    # Control Voltage check
    if len(semiMachine.controlVEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Control voltage exceeds given limit of 15 characters")
        return

    # Total motor check
    if len(semiMachine.totMotorEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Total Motor entry exceeds given limit of 15 characters")
        return

    # Full load check
    if len(semiMachine.fullLoadEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Full load entry exceeds given limit of 15 characters")
        return 

    # Engineer initials check
    if len(semiMachine.enginEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Engineer initials exceeds the given limit of 15 characters")
        return

    # Date variable and check - format ##/##/####
    if len(semiMachine.dateEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Date exceeds the given limit of 15 characters")
        return
    reDateText = re.search("(?:0[1-9]|1[012])/(?:0[1-9]|[12][0-9]|3[01])/(?:19\d{2}|20[0-9][0-9])", semiMachine.dateEntry.get())
    if not reDateText:
        messagebox.showerror(title="Invalid Entry", message="Date format is MM/DD/YYYY")
        return

    # Folder variables
    sourceFolder = "G:\\~~EE Directory~~\\01-AutoCAD Electrical\\07-Standard Drawings Templates\\Synergy Semi Auto"
    destinationFolder = os.getcwd() + "\\" + semiMachine.projNumEntry.get() + " Drawings"

    # new folder destination
    try: 
        os.mkdir(destinationFolder)
    except:
        res = messagebox.askyesno(title="Job Already Exists", message="Job already exists in directory, replace?")
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
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + semiMachine.projNumEntry.get() + " Drawings.aepx")
        elif files[1] == "wdp":
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + semiMachine.projNumEntry.get() + " Drawings.wdp")
        elif files[1] == "dwg":
            noNumFile = fileName.split("_")
            noNumFileNoEx = noNumFile[1].split(".")
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + semiMachine.projNumEntry.get() + "_" + noNumFileNoEx[0] + ".dwg")
        elif files[1] == "bak":
            noNumFile = fileName.split("_")
            noNumFileNoEx = noNumFile[1].split(".")
            os.rename(destinationFolder + "\\" + fileName, destinationFolder + "\\" + semiMachine.projNumEntry.get() + "_" + noNumFileNoEx[0] + ".bak")
        else: pass

    # setting running to true for drawings edit script
    drawingsEditLayers.running = True

    # creation of edit thread to integrate stop button
    global editThread
    editThread = threading.Thread(target=doneThreadFunc, args=(destinationFolder, semiMachine))
    
    # check to see if program has been stopped or not
    if semiMachine.numRan == 0:
        editThread.start()
        semiMachine.numRan += 1
    else:
        editThread = threading.Thread(target=doneThreadFunc, args=(destinationFolder, semiMachine))
        editThread.start()

##########################################################
# Threading function for autoDoneCallBack to help stop   #
##########################################################
def autoDoneThreadFunc(destFolder, autoMachine):
    # Calling auto drawings edit method to update drawing to user input
    autoDrawingsEditLayers.editLayers(destFolder, autoMachine.custEntry.get(), autoMachine.distrEntry.get(), autoMachine.projNumEntry.get(), autoMachine.manYearEntry.get(), autoMachine.phaseEntry.get(),
                        autoMachine.mainLineVEntry.get(), autoMachine.controlVEntry.get(), autoMachine.totMotorEntry.get(), autoMachine.fullLoadEntry.get(), autoMachine.enginEntry.get(), autoMachine.dateEntry.get(),
                        autoMachine.autoSynVar.get(), autoMachine.motorsEVar.get(), autoMachine.motorsXVar.get(), autoMachine.entryDrives.get(), autoMachine.exitDrives.get(), autoMachine.autoULVar.get())

##########################################################################################################
# method that is called when user has entered all fields and clicks done button for automatic machines   #
##########################################################################################################
def autoDoneCallBack(autoMachine):
    # variables and checks
    # Customer variable and check
    if len(autoMachine.custEntry.get()) > 25:
        messagebox.showerror(title="Invalid Entry", message="Customer name exceeds given limit of 25 characters.")
        return

    # Distributor variable and check
    if len(autoMachine.distrEntry.get()) > 25:
        messagebox.showerror(title="Invalid Entry", message="Distributor name exceeds given limit of 25 characters.")
        return

    # Project number variable and check
    if len(autoMachine.projNumEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Project number exceeds given limit of 15 characters.")
        return

    # Manufacturer year variable and check
    if len(autoMachine.manYearEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Manufacturing year exceeds given limit of 15 characters.")
        return

    # Phase variable and check
    if len(autoMachine.phaseEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Phase entry exceeds given limit of 15 characters.")
        return
    
    # Main line voltage variable and check
    if len(autoMachine.mainLineVEntry.get()) > 18:
        messagebox.showerror(title="Invalid Entry", message="Main line voltage exceeds given limit of 18 characters.")
        return

    # Control voltage variable and check
    if len(autoMachine.controlVEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Control voltage exceeds given limit of 15 characters.")
        return

    # Total motor variable and check
    if len(autoMachine.totMotorEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Total motor entry exceeds given limit of 15 characters.")
        return

    # Full load variable and check
    if len(autoMachine.fullLoadEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Full load entry exceeds given limit of 15 characters.")
        return
    
    # Engineer initials varaible and check
    if len(autoMachine.enginEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Engineer initials exceeds given limit of 15 characters.")
        return

    # Date variable and check - foramt ##/##/####
    if len(autoMachine.dateEntry.get()) > 15:
        messagebox.showerror(title="Invalid Entry", message="Date exceeds given limit of 15 characters.")
        return
    reDateText = re.search("(?:0[1-9]|1[012])/(?:0[1-9]|[12][0-9]|3[01])/(?:19\d{2}|20[0-9][0-9])", autoMachine.dateEntry.get())
    if not reDateText:
        messagebox.showerror(title="Invalid Entry", message="Date format is MM/DD/YYYY")
        return

    # folder variables
    sourceFolder = "G:\\~~EE Directory~~\\01-AutoCAD Electrical\\07-Standard Drawings Templates\\Synergy Auto"
    destinationFolder = os.getcwd() + "\\" + autoMachine.projNumEntry.get() + " Drawings"

    # new folder destination
    try:
        os.mkdir(destinationFolder)
    except:
        res = messagebox.askyesno(title="Job Already Exits", message="Job already exists in directory, replace?")
        if res == True:
            shutil.rmtree(destinationFolder)
            os.mkdir(destinationFolder)
        else: return

    # Copies original schematic to new folder
    source = sourceFolder + "\\800500 - Synergy 5 Auto Schematics.dwg"
    destination = destinationFolder + "\\800500 - Synergy 5 Auto Schematics.dwg"
    if os.path.isfile(source):
        shutil.copy(source, destination)

    # Renames new folders file to match user inputs
    os.rename(destinationFolder + "\\800500 - Synergy 5 Auto Schematics.dwg", destinationFolder + "\\" + autoMachine.projNumEntry.get() + " Schematics.dwg")
    
    # Setting running to true for auto drawings edit script
    autoDrawingsEditLayers.running = True

    # Creation of edit thread to integrate stop button
    global autoEditThread
    autoEditThread = threading.Thread(target=autoDoneThreadFunc, args=(destinationFolder, autoMachine))

    # Check to see if program has been stopped or not
    if autoMachine.numRan == 0:
        autoEditThread.start()
        autoMachine.numRan += 1
    else:
        autoEditThread = threading.Thread(target=autoDoneThreadFunc, args=(destinationFolder, autoMachine))
        autoEditThread.start()

#################################################################
# Method called when user presses stop button on main tab       #
#################################################################
def stopCallBack(projNumEntry, macType):
    if macType == 0:
        drawingsEditLayers.running = False
        messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopped editing " + projNumEntry.get() + " Drawings.")
        editThread.join(3)
    else:
        autoDrawingsEditLayers.running = False
        messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopeed editing " + projNumEntry.get() + " Drawings.")
        autoEditThread.join(3)

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
        messagebox.showerror(title="Job Number Empty", message="Job number field empty, please enter job number")
        return
    
    # Check to confirm manual directory is not empty when using manual directory
    if mDir == "" and exJobNumEntry.get() == 1:
        messagebox.showerror(title="Manual Directory Empty", message="Manual directory field empty, please enter manual directory")
        return

    # Check to make sure automatic directory exists
    testFile = dir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(testFile) == False and exJobNumEntry.get() == 0:
        messagebox.showerror(title="Invalid Directory", message="Automatic finder could not find job number directory, please enter manual directory")
        return

    # Check to make sure manual directory exists
    mTestFile = mDir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(mTestFile) == False and exJobNumEntry.get() == 1:
        messagebox.showerror(title="Invalid Directory", message="Manual directory is invalid, please try again")
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
    messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopped exporting " + exJobNumEntry.get() + " Drawings.")

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
        messagebox.showerror(title="Job Number Empty", message="Job number field empty, please enter job number")
        return

    # Check to confirm manual directory is not empty when using manual directory
    if mDir == "" and accessJobVar.get() == 1:
        messagebox.showerror(title="Manual Directory Empty", message="Maunal directory field empty, please enter manual directory")
        return
    
    # Check to make sure automatic directory exists
    testFile = dir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(testFile) == False and accessJobVar.get() == 0:
        messagebox.showerror(title="Invalid Directory", message="Automtic finder could not find job number directory, please enter maunal directory")
        return
    
    # Check to make sure maunal directory exists
    if os.path.exists(mDir) == False and accessJobVar.get() == 1:
        messagebox.showerror(title="Invalid Directory", message="Manual directory is invalid, please try again")
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
        messagebox.showerror(title="Error", message="This error should not happen")
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
        messagebox.showerror(title="No Printer Selected", message="Please select a printer")
        return
    else:
        printer = printerVar.get()

    #
    # general error proof checks
    #
    # checks for blank text boxes
    # jobNum can never be empty
    if jobNum == "":
        messagebox.showerror(title="Job Number Empty", message="Job number field empty, please enter job number")
        return
    
    # check to confirm manual directory is not empty when using manual search
    if mDir == "" and jobVar.get() == 1:
        messagebox.showerror(title="Manual Directory Empty", message="Manual directory field empty, please enter manual directory")
        return
    
    # check to confirm pages is not empty when printing selected pages
    if pages == "" and paVar.get() == 0:
        messagebox.showerror(title="Pages Empty", message="Pages field empty, please enter desired pages to print")
        return

    # check to confirm pages is only integers
    tempPages = pages
    tempPages = tempPages.replace(',', '')
    tempPages = tempPages.replace('-', '')
    if tempPages.isnumeric() == False and paVar.get() == 0:
        messagebox.showerror(title="Pages Not Numbers", message="Pages field contains invalid characters, please try again")
        return

    # check to confirm jobNum is valid
    jobCheck = int(jobNum) // 10
    if dir == "Maybe":
        if jobCheck.isnumeric() == False:
            messagebox.showerror(title="Job Number Not Numbers", message="Job number field contains invalid characters, please try again")
            return
    
    # check to confirm job number directory exists
    testFile = dir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(testFile) == False and jobVar.get() == 0:
        messagebox.showerror(title="Invalid Directory", message="Automatic finder could not find job number directory, please enter manual directory")
        return

    # check to confirm manual directory exists
    mTestFile = mDir + "\\" + jobNum + "_Sheet00.dwg"
    if os.path.isfile(mTestFile) == False and jobVar.get() == 1:
        messagebox.showerror(title="Invalid Directory", message="Manual directory is invalid, please try again")
        return

    # check to confirm printer is not empty
    if printer == "":
        messagebox.showerror(title="Printer Empty", message="Printer field is empty, please enter printer name")
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
    messagebox.showinfo(title="Program Stopped", message="Stop button was pushed and the program stopped printing " + jobEntry.get() + " Drawings.")
    printThread.join(3)


