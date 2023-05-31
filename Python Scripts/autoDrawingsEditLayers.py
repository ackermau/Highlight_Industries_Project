from pyautocad import Autocad
from tkinter import *
from tkinter import messagebox
import os
import psutil
import time
import tkinter as tk

# global variables
running = True

# Functions to edit drawings by layers
def editLayers(folder, cust, distr, projNum, manYear, phase, mainLineV, controlV, totMotor, 
            fullLoad, engin, date, autoSynVar, motorsEVar, motorsXVar, entryDrives, exitDrives, autoULVar, forkLiftVar, wzLiftVar):
    scriptDir = os.getcwd() + "\\AutoCAD Script\\Synergy Automatic Scripts"

    ##########################
    # Edit auto schematics   #
    ##########################
    drawing = projNum + " Schematics.dwg"
    dir = folder + "\\" + drawing
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            while True:
                try:
                    acad.doc.SendCommand('FILEDIA 0 \n')
                    break
                except: pass
            
            ##################################
            # Automatic page 0 edit script   #
            ##################################
            # Edit Script 
            if running == True:
                with open(scriptDir + '\\syn5AutoP0EditScript.scr', 'r') as file:
                    script = file.read()
                    script = script.replace('<CUSTOMER>', cust)
                    script = script.replace('<DISTRIBUTOR>', distr)
                    script = script.replace('<PROJNUM>', projNum)
                    script = script.replace('<MANYEAR>', manYear)
                    script = script.replace('<PHASE>', phase)
                    script = script.replace('<MAINLINE>', mainLineV)
                    script = script.replace('<CONTROLV>', controlV)
                    script = script.replace('<TOTMOTOR>', totMotor)
                    script = script.replace('<FULLLOAD>', fullLoad)
                with open(scriptDir + '\\tempP0Script.scr', 'w') as file:
                    file.write(script)
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60: return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempP0Script.scr \n')
                        break
                    except: pass
            else: return
            
            # Entry Scripts
            if running == True:
                if motorsEVar == 1:
                    for i in range(1, 6):
                        if entryDrives >= i:
                            with open(scriptDir + '\\syn5AutoP0En' + str(i) + '.scr', 'r') as file:
                                script = file.read()
                                script = script.replace('<PROJNUM>', projNum)
                            with open(scriptDir + '\\tempP0Script.scr', 'w') as file:
                                file.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempP0Script \n')
                                    break
                                except: pass
            else: return

            # Exit Scripts
            if running == True:
                if motorsXVar == 1:
                    for i in range(1, 6):
                        if exitDrives >= i:
                            with open(scriptDir + '\\syn5AutoP0Ex' + str(i) + '.scr', 'r') as file:
                                script = file.read()
                                script = script.replace('<PROJNUM>', projNum)
                            with open(scriptDir + '\\tempP0Script.scr', 'w') as file:
                                file.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempP0Script \n')
                                    break
                                except: pass
            else: return

            #############################
            # Synergy 5 border script   #
            #############################    
            if running == True:
                with open(scriptDir + '\\syn5AutoBorderScript.scr', 'r') as file:
                    script = file.read()
                    script = script.replace('<ENGINIT>', engin)
                    script = script.replace('<DATE>', date)
                    script = script.replace('<PROJNUM>', projNum)
                with open(scriptDir + '\\tempBorderScript.scr', 'w') as file:
                    file.write(script)
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60: return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempBorderScript.scr \n')
                        break
                    except: pass
            else: return

            ########################
            # Layer manipulation   #
            ########################
            # Entry drives
            if running == True:
                if motorsEVar == 1:
                    for drive in list(range(1, entryDrives + 1)):
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T ENTRY' + str(drive) + ' \n\n')
                                break
                            except: pass
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER F NOENTRY' + str(drive) + ' \n\n')
                                break
                            except: pass

                    if entryDrives < 5:
                        for drive in list(range(entryDrives + 1, 6)):
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER F ENTRY' + str(drive) + ' \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER T NOENTRY' + str(drive) + ' \n\n')
                                    break
                                except: pass
                else:
                    for drive in range(1, 6):
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER F ENTRY' + str(drive) + ' \n\n')
                                break
                            except: pass
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T NOENTRY' + str(drive) + '\n\n')
                                break
                            except: pass
            else: return
            
            # Exit drives
            if running == True:
                if motorsXVar == 1:
                    for drive in list(range(1, exitDrives + 1)):
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T EXIT' + str(drive) + ' \n\n')
                                break
                            except: pass
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER F NOEXIT' + str(drive) + ' \n\n')
                                break
                            except: pass
      
                    if exitDrives < 5:
                        for drive in list(range(exitDrives + 1, 6)):
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER F EXIT' + str(drive) + ' \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER T NOEXIT' + str(drive) + ' \n\n')
                                    break
                                except: pass
                else:
                    for drive in range(1, 6):
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER F EXIT' + str(drive) + ' \n\n')
                                break
                            except: pass
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T NOEXIT' + str(drive) + ' \n\n')
                                break
                            except: pass
            else: return 

            # End of Exit layers
            if running == True:
                if motorsXVar == 1:
                    # Turning off all end exit layers
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F ENDEXIT1 \n\n')
                            break
                        except: pass
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F ENDEXIT2 \n\n')
                            break
                        except: pass
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F ENDEXIT3 \n\n')
                            break
                        except: pass
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F ENDEXIT4 \n\n')
                            break
                        except: pass
                    # Turning on specific end exit layer
                    if exitDrives == 1:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T ENDEXIT1 \n\n')
                                break
                            except: pass
                    elif exitDrives == 2:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T ENDEXIT2 \n\n')
                                break
                            except: pass
                    elif exitDrives == 3:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T ENDEXIT3 \n\n')
                                break
                            except: pass
                    elif exitDrives == 4:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('-LAYER T ENDEXIT4 \n\n')
                                break
                            except: pass

            # UL Option
            if running == True:
                if autoULVar == 1:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER T UL \n\n')
                            break
                        except: pass
                else:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F UL \n\n')
                            break
                        except: pass

            # Forklift Option
            if running == True:
                if forkLiftVar == 1:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER T FORKLIFT \n\n')
                            break
                        except: pass
                else:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F FORKLIFT \n\n')
                            break
                        except: pass

            # Wrap Zone Lift Option
            if running == True:
                if wzLiftVar == 1:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER T WZLIFT \n\n')
                            break
                        except: pass
                else:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F WZLIFT \n\n')
                            break
                        except: pass

            ################################
            # Freedom or Synergy scripts   #
            ################################
            if autoSynVar == 0:
                if running == True:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER T SYN5 \n\n')
                            break
                        except: pass
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F FREEDOM \n\n')
                            break
                        except: pass
                else: return
            else:
                if running == True:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER F SYN5 \n\n')
                            break
                        except: pass
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('-LAYER T FREEDOM \n\n')
                            break
                        except: pass
                else: return

            ################
            # End script   #
            ################
            if running == True:
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60: return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\endScript.scr \n')
                        break
                    except: pass
        # AutoCAD not running error
        else: return
    else: return

    # message sent to user telling them the drawing has been edited for jobNum
    tk.messagebox.showinfo(title="Finished Editing Drawing", message="Job " + projNum + " has been created in '" + projNum + " Drawings' folder")