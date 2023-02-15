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
def editLayers(folder, cust, distr, projNum, manYear, phase, mainLine, control, totMotor, 
            fullLoad, engin, date, syn, profile, scale, cold, AMP20, split, auto, ul, door, har):
    scriptDir = os.getcwd() + "\\AutoCAD Script\\Synergy Semi Auto Scripts"

    if syn == 0:
        synType = "SYNERGY 2"
        if profile == 0:
            profType = "HIGH PROFILE"
            synProf = "SYNERGY 2 HP"
        else:
            profType = "LOW PROFILE"
            synProf = "SYNERGY 2 LP"
    elif syn == 1:
        synType = "SYNERGY 2.5"
        if profile == 0:
            profType = "HIGH PROFILE"
            synProf = "SYNERGY 2.5 HP"
        else:
            profType = "LOW PROFILE"
            synProf = "SYNERGY 2.5 HP"
    elif syn == 2:
        synType = "SYNERGY 3"
        if profile == 0:
            profType = "HIGH PROFILE"
            synProf = "SYNERGY 3 HP"
        else:
            profType = "LOW PROFILE"
            synProf = "SYNERGY 3 LP"
    else:
        synType = "SYNERGY 4"
        if profile == 0:
            profType = "HIGH PROFILE"
            synProf = "SYNERGY 4 HP"
        else:
            profType = "LOW PROFILE"
            synProf = "SYNERGY 4 LP"

    # Wdp file
    wdpFile = folder + "\\" + projNum + " Drawings.wdp"
    with open(wdpFile, 'r') as file:
        script = file.read() 
        script = script.replace('SynergySemiAuto', projNum)
        script = script.replace('XXXX', projNum)
        script = script.replace('X.X.X.', engin)
        script = script.replace('XX/XX/XXXX', date)
    with open(wdpFile, 'w') as file:
        file.write(script)

    ################
    # Drawing 00   #
    ################
    sheet = projNum + "_Sheet00.dwg"
    dir = folder + "\\" + sheet
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            while True:
                try:
                    acad.doc.SendCommand('FILEDIA 0 \n')
                    break
                except: pass

            if running == True:
                # Synergy type script
                if syn == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-syn2Script.scr \n')
                            break
                        except: pass
                else:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60: return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-otherSynScript.scr \n')
                            break
                        except: pass
            else: return

            if running == True:
                # Scale and auto film scripts
                if scale == 1:
                    if auto == 1:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-scaleAutoScript.scr \n')
                                break
                            except: pass
                    else:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-scaleScript.scr \n')
                                break
                            except: pass
                else:
                    if auto == 1:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-autoFilmScript.scr \n')
                                break
                            except: pass
                    else:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60: return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-noSnoAScript.scr \n')
                                break
                            except: pass
            else: return

            if running == True:
                # Machine type script
                with open(scriptDir + '\\00-macTypeScript.scr', 'r') as file:
                    script = file.read()
                    script = script.replace('<SYNERGYTYPE>', synType)
                with open(scriptDir + '\\tempScript.scr', 'w') as file:
                    file.write(script)
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60: return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except: pass
            else: return

            if running == True:
                # Profile type script
                with open(scriptDir + '\\00-profileScript.scr', 'r') as file:
                    script = file.read()
                    script = script.replace('<PROFILETYPE>', profType)
                with open(scriptDir + '\\tempScript.scr', 'w') as file:
                    file.write(script)
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60: return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except: pass
            else: return

            if running == True:
                # Edit script
                with open(scriptDir + '\\00-EditScript.scr', 'r') as file:
                    script = file.read()
                    script = script.replace('<CUSTOMER>', cust)
                    script = script.replace('<DISTRIBUTOR>', distr)
                    script = script.replace('<PNUM>', projNum)
                    script = script.replace('<MANYEAR>', manYear)
                    script = script.replace('<PHASE>', phase)
                    script = script.replace('<MAINLINE>', mainLine)
                    script = script.replace('<CONTROLV>', control)
                    script = script.replace('<TOTMOTOR>', totMotor)
                    script = script.replace('<FULLLOAD>', fullLoad)
                    script = script.replace('<ENGINIT>', engin)
                    script = script.replace('<DATE>', date)
                    script = script.replace('<SYNPROF>', synProf)
                with open(scriptDir + '\\tempScript.scr', 'w') as file:
                    file.write(script)
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60: return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except: pass
            else: return

        # autocad not running error
        else: return
    else: return

    if running == True:
        ######################
        # Drawings 01 - 32   #
        ######################
        for x in range(1, 33):
            if running == True:
                if x < 10:
                    dir = folder + "\\" + projNum + "_Sheet0" + str(x) + ".dwg"
                else:
                    dir = folder + "\\" + projNum + "_Sheet" + str(x) + ".dwg"
                os.startfile(dir)
                if "acad.exe" in (i.name() for i in psutil.process_iter()):
                    acad = Autocad()

                    ###############################
                    # AMP layers edit - p.01-32   #
                    ###############################
                    if running == True:
                        if cold == 1 or AMP20 == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON 20AMP \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF 15AMP \n\n')
                                    break
                                except: pass
                        else:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF 20AMP \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON 15AMP \n\n')
                                    break
                                except: pass

                    ################################
                    # Cold layers edit - p.01-32   #
                    ################################
                    # Cold option is turned off
                    if running == True:
                        if cold == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter() 
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF COLD \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOCOLD \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF COLDSYN2 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if(endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF COLDSYN2_5 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF COLDSYN4 \n\n')
                                    break
                                except: pass
                    else: return
                        
                    # Cold option is turned on
                    if running == True:
                        if cold == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON COLD \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()    
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF NOCOLD \n\n')
                                    break
                                except: pass
                            
                            # Page 1 coldsyn scripts
                            if x == 1:
                                if cold == 1:
                                    # Synergy 3
                                    if syn == 2:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER ON COLDSYN2 \n\n')
                                                break
                                            except: pass
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if(endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER ON COLDSYN2_5 \n\n')
                                                break
                                            except: pass
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER OFF COLDSYN4 \n\n')
                                                break
                                            except: pass
                                    else:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER OFF COLDSYN2 \n\n')
                                                break
                                            except: pass
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if(endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER OFF COLDSYN2_5 \n\n')
                                                break
                                            except: pass
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER OFF COLDSYN4 \n\n')
                                                break
                                            except: pass
                                    # Synergy 2
                                    if syn == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER ON COLDSYN2 \n\n')
                                                break
                                            except: pass
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER OFF COLDSYN4 \n\n')
                                                break
                                            except: pass

                                    # Synergy 2.5
                                    if syn == 1:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER ON COLDSYN2_5 \n\n')
                                                break
                                            except: pass
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER OFF COLDSYN4 \n\n')
                                                break
                                            except: pass

                                    # Synergy 4
                                    if syn == 3:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('-LAYER ON COLDSYN4 \n\n')
                                                break
                                            except: pass               
                    else: return

                    ################################
                    # Scale layer edit - p.01-32   #
                    ################################
                    # Scale option is turned off
                    if running == True:
                        if scale == 0:
                            if x == 1:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('-LAYER OFF 15AMPS \n\n')
                                        break
                                    except: pass
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('-LAYER OFF 20AMPS \n\n')
                                        break
                                    except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SCALE \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOSCALE \n\n')
                                    break
                                except: pass
                    else: return
                    
                    # Scale option is turned on
                    if running == True:
                        if scale == 1:
                            if x == 1:
                                if cold == 1 or syn == 3:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER ON 20AMPS \n\n')
                                            break
                                        except: pass
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER OFF 15AMPS \n\n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER OFF 20AMPS \n\n')
                                            break
                                        except: pass
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER ON 15AMPS \n\n')
                                            break
                                        except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SCALE \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF NOSCALE \n\n')
                                    break
                                except: pass
                    else: return

                    ######################################
                    # AutoFilmCut layer edit - p.01-32   #
                    ######################################
                    # AutoFilmCut option is turned off
                    if running == True:
                        if auto == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF AUTOFILMCUT \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOAUTOFILMCUT \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF AUTOUL \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOAUTOUL \n\n')
                                    break
                                except: pass
                    else: return
                    
                    # AutoFilmCut option is turned on
                    if running == True:
                        if auto == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON AUTOFILMCUT \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF NOAUTOFILMCUT \n\n')
                                    break
                                except: pass

                            # Page 3 AutoUL scripts
                            if x == 3:
                                if ul == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER OFF AUTOUL \n\n')
                                            break
                                        except: pass
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER ON NOAUTOUL \n\n')
                                            break
                                        except: pass
                                if ul == 1:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER ON AUTOUL \n\n')
                                            break
                                        except: pass
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER OFF NOAUTOUL \n\n')
                                            break
                                        except: pass
                    else: return
                    
                    #####################################
                    # SplitFrame layer edit - p.01-32   #
                    #####################################
                    # SplitFrame option is turned off
                    if running == True:
                        if split == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SPLITFRAME \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOSPLITFRAME \n\n')
                                    break
                                except: pass
                    else: return

                    # SplitFrame option is turned on
                    if running == True:
                        if split == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SPLITFRAME \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF NOSPLITFRAME \n\n')
                                    break
                                except: pass
                    else: return
                    
                    #############################
                    # UL layer edit - p.01-32   #
                    #############################
                    # UL option is turned off
                    if running == True:
                        if ul == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF UL \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOUL \n\n')
                                    break
                                except: pass
                    else: return

                    # UL option is turned on
                    if running == True:
                        if ul == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON UL \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF NOUL \n\n')
                                    break
                                except: pass
                    else: return

                    #######################################
                    # CarriageDoor layer edit - p.01-32   #
                    #######################################
                    # CarriageDoor option is turned off 
                    if running == True:
                        if door == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF CARRIAGEDOOR \n\n')
                                    break
                                except: pass
                    else: return

                    # CarriageDoor option is turned on
                    if running == True:
                        if door == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON CARRIAGEDOOR \n\n')
                                    break
                                except: pass
                    else: return

                    ###############################################
                    # Synergy 2, 2.5, 3, 4 layer edit - p.01-32   #
                    ###############################################
                    # Synergy 2 option is turned on
                    if running == True:
                        if syn == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SYN2 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SYN2_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2_5_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN4 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOSYN4 \n\n')
                                    break
                                except: pass
                            if x == 15:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('-LAYER OFF HAR \n\n')
                                        break
                                    except: pass
                    else: return
                    
                    # Synergy 2.5 option is turned on
                    if running == True:
                        if syn == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SYN2_5_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN4 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOSYN4 \n\n')
                                    break
                                except: pass
                            if x == 15:
                                if har == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER OFF HAR \n\n')
                                            break
                                        except: pass
                                if har == 1:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER ON HAR \n\n')
                                            break
                                        except: pass
                    else: return

                    # Synergy 3 option is turned on
                    if running == True:
                        if syn == 2:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SYN2_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SYN2_5_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN4 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON NOSYN4 \n\n')
                                    break
                                except: pass
                            if x == 15:
                                if har == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER OFF HAR \n\n')
                                            break
                                        except: pass
                                if har == 1:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('-LAYER ON HAR \n\n')
                                            break
                                        except: pass
                    else: return

                    # Synergy 4 option is turned on
                    if running == True:
                        if syn == 3:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER ON SYN4 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF NOSYN4 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2_5_3 \n\n')
                                    break
                                except: pass
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('-LAYER OFF SYN2 \n\n')
                                    break
                                except: pass
                            if x == 15:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('-LAYER OFF HAR \n\n')
                                        break
                                    except: pass
                    else: return
                        

                    ########################
                    # Table scripts - 25   #
                    ########################
                    if running == True:
                        if x == 25:
                            if running == True:
                                # Synergy scripts
                                if syn == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-syn2Script.scr \n')
                                            break
                                        except: pass
                                elif syn == 1:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-syn2_5Script.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-syn3Script.scr \n')
                                            break
                                        except: pass
                            else: return
                            if running == True:
                                # Carriage door scripts
                                if door == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noCarriageDoorScript.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-carriageDoorScript.scr \n')
                                            break
                                        except: pass
                            else: return
                            if running == True:
                                # HAR scripts
                                if har == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noHARScript.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-HARScript.scr \n')
                                            break
                                        except: pass
                            else: return
                            if running == True:
                                # Cold and auto film scripts
                                if auto == 1 and cold == 1:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-coldAutoFilmScript \n')
                                            break
                                        except: pass
                                elif auto == 0 and cold == 1:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-coldNoAutoScript.scr \n')
                                            break
                                        except: pass
                                elif auto == 1 and cold == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noColdAutoScript.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noColdNoAutoFilmScript.scr \n')
                                            break
                                        except: pass
                            else: return
                    else: return

                    ########################
                    # Table scripts - 29   #
                    ########################
                    if running == True:
                        if x == 29:
                            # Cold package scripts
                            if cold == 0:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noColdScript.scr \n')
                                        break
                                    except: pass
                                # Cold + Synergy scripts
                                if syn == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noColdSyn2Script.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noColdOtherScript.scr \n')
                                            break
                                        except: pass
                            else:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-coldScript.scr \n')
                                        break
                                    except: pass

                            if running == True:
                                # Scale package script
                                if scale == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noScaleScript.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-scaleScript.scr \n')
                                            break
                                        except: pass
                            else: return
                            
                            if running == True:
                                # Auto film scripts
                                if auto == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noAutoFilmScript.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-autoFilmScript.scr \n')
                                            break
                                        except: pass
                                    # Auto + Cold scripts
                                    if cold == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-autoNoColdScript.scr \n')
                                                break
                                            except: pass
                                    else:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-autoColdScript.scr \n')
                                                break
                                            except: pass
                            else: return

                            if running == True:
                                # UL script
                                if ul == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULScript.scr \n')
                                            break
                                        except: pass
                                    # UL + Scale script
                                    if scale == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULNoScaleScript.scr \n')
                                                break
                                            except: pass
                                    else:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULScaleScript.scr \n')
                                                break
                                            except: pass
                                    # UL + Auto script
                                    if auto == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULNoAutoScript.scr \n')
                                                break
                                            except: pass
                                    else:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULAutoScript.scr \n')
                                                break
                                            except: pass
                                else:
                                    if auto == 1:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-ULAutoScript.scr \n')
                                                break
                                            except: pass
                            else: return

                            if running == True:
                                # Synergy scripts
                                if syn == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-syn2Script.scr \n')
                                            break
                                        except: pass
                                    # HAR + Carriage door scripts with synergy 2
                                    if har == 0 and door == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-0Script.scr \n')
                                                break
                                            except: pass
                                    elif (har == 1 and door == 0) or (har == 0 and door == 1):
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-1Script.scr \n')
                                                break
                                            except: pass
                                    else: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-otherSynScript.scr \n')
                                            break
                                        except: pass
                                    # HAR + Carriage door scripts with synergy 2.5 and 3
                                    if har == 0 and door == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-1Script.scr \n')
                                                break
                                            except: pass
                                    elif (har == 1 and door == 0) or (har == 0 and door == 1):
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-2Script.scr \n')
                                                break
                                            except: pass
                                    else: pass
                            else: return
                    else: return

                    ########################
                    # Table scripts - 32   #
                    ########################
                    if running == True:
                        if x == 32:
                            # Synergy scripts
                            if syn == 0:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-syn2Script.scr \n')
                                        break
                                    except: pass
                            else:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-otherSynScript.scr \n')
                                        break
                                    except: pass
                            
                            if running == True:
                                # Cold package scripts
                                if cold == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-noColdScript.scr \n')
                                            break
                                        except: pass
                                    # Synergy + Cold scripts
                                    if syn == 0:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-syn2NoColdScript.scr \n')
                                                break
                                            except: pass
                                    else:
                                        timer = time.perf_counter()
                                        while True:
                                            endTimer = time.perf_counter()
                                            if (endTimer - timer) >= 60: return
                                            try:
                                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-synONoColdScript.scr \n')
                                                break
                                            except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-coldScript.scr \n')
                                            break
                                        except: pass
                            else: return
                            
                            if running == True:
                                # Split frame scripts
                                if split == 0:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-noSplitFrameScript.scr \n')
                                            break
                                        except: pass
                                else:
                                    timer = time.perf_counter()
                                    while True:
                                        endTimer = time.perf_counter()
                                        if (endTimer - timer) >= 60: return
                                        try:
                                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-splitFrameScript.scr \n')
                                            break
                                        except: pass
                            else: return
                    else: return

                    ########################
                    # Table scripts - 30   #
                    ########################
                    if running == True:
                        if x == 30:
                           # Synergy script
                            if syn == 0:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\30-syn2Script.scr \n')
                                        break
                                    except: pass
                            else:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60: return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\30-otherSynScript.scr \n')
                                        break
                                    except: pass 
                    else: return

                    ############################
                    # Edit scripts - 01 - 27   #
                    ############################
                    if running == True:
                        if x in range(1, 28):
                            with open(scriptDir + '\\01-27-EditScript.scr', 'r') as file:
                                script = file.read()
                                script = script.replace('<PNUM>', projNum)
                                script = script.replace('<ENGINIT>', engin)
                                script = script.replace('<DATE>', date)
                                script = script.replace('<SYNPROF>', synProf)
                            with open(scriptDir + '\\tempScript.scr', 'w') as file:
                                file.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                                    break
                                except: pass
                    else: return
                    
                    ######################
                    # Edit script - 28   #
                    ######################
                    if running == True:
                        if x == 28:
                            with open(scriptDir + '\\28-EditScript.scr', 'r') as file:
                                script = file.read()
                                script = script.replace('<PNUM>', projNum)
                                script = script.replace('<ENGINIT>', engin)
                                script = script.replace('<DATE>', date)
                                script = script.replace('<SYNPROF>', synProf)
                            with open(scriptDir + '\\tempScript.scr', 'w') as file:
                                file.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                                    break
                                except: pass
                    else: return

                    ######################
                    # Edit script - 30   #
                    ######################
                    if running == True:
                        if x == 30:
                            with open(scriptDir + '\\30-EditScript.scr', 'r') as file:
                                script = file.read()
                                script = script.replace('<PNUM>', projNum)
                                script = script.replace('<ENGINIT>', engin)
                                script = script.replace('<DATE>', date)
                                script = script.replace('<SYNPROF>', synProf)
                            with open(scriptDir + '\\tempScript.scr', 'w') as file:
                                file.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                                    break
                                except: pass
                    else: return

                    ##############################
                    # Edit script - 29, 31, 32   #
                    ##############################
                    if running == True:
                        if x == 29 or x == 31 or x == 32:
                            with open(scriptDir + '\\29_31_32-EditScript.scr', 'r') as file:
                                script = file.read()
                                script = script.replace('<PNUM>', projNum)
                                script = script.replace('<ENGINIT>', engin)
                                script = script.replace('<DATE>', date)
                                script = script.replace('<SYNPROF>', synProf)
                            with open(scriptDir + '\\tempScript.scr', 'w') as file:
                                file.write(script)
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60: return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                                    break
                                except: pass
                    else: return

                # AutoCad not running error
                else: return        
            else: return
    else: return
    
    sheet = projNum + "_Sheet00.dwg"
    dir = folder + "\\" + sheet
    if running == True:
        ################
        # End script   #
        ################
        os.startfile(dir)
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
        else: return
    else: return

    # message sent to user telling them that the drawings have been edited for jobNum
    tk.messagebox.showinfo(title="Finished Editing Drawings", message="Job " + projNum + " has been created in '" + projNum + " Drawings' folder")
                        