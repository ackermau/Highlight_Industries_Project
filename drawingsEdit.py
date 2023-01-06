from pyautocad import Autocad
from tkinter import *
from tkinter import messagebox
import os
import psutil
import time
import tkinter as tk

# global variables
running = True

# Edits drawings to update to user inputs
def editDrawings(folder, cust, distr, projNum, manYear, phase, mainLine, control, totMotor, fullLoad, engin, date, syn, profile, scale, cold, split, photo, auto, ul, door, har):
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
    else:
        synType = "SYNERGY 3"
        if profile == 0:
            profType = "HIGH PROFILE"
            synProf = "SYNERGY 3 HP"
        else:
            profType = "LOW PROFILE"
            synProf = "SYNERGY 3 LP"

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


    # Drawing 00
    sheet = projNum + "_Sheet00.dwg"
    dir = folder + "\\" + sheet
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            if running == True:
                # Synergy type script
                if syn == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-syn2Script.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-otherSynScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Scale and auto film scripts
                if scale == 1:
                    if auto == 1:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-scaleAutoScript.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-scaleScript.scr \n')
                                break
                            except:
                                pass
                else:
                    if auto == 1:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-autoFilmScript.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\00-noSnoAScript.scr \n')
                                break
                            except:
                                pass
            else:
                return

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
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return

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
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return

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
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return

        # autocad not running error
        else: 
            return
    else:
        return

    # Drawing 01
    dir = folder + "\\" + projNum + "_Sheet01.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
        
            if running == True:
                # Cold package script
                if cold == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-noColdScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-coldScript.scr \n')
                            break
                        except:
                            pass
                    # Synergy scripts inside cold option
                    if syn == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-syn2Script.scr \n')
                                break
                            except:
                                pass
                    elif syn == 1:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-syn2_5Script.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-syn3Script.scr \n')
                                break
                            except:
                                pass
            else:
                return

            if running == True:
                # Scale package script
                if scale == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-noScaleScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-scaleScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Auto film script
                if auto == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-noAutoFilmScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\01-autoFilmScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Edit script
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
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return

        # autocad not running error
        else:
            return
    else:
        return

    if running == True:
        # Drawings 02 - 09
        for x in range(2, 10):
            dir = folder + "\\" + projNum + "_Sheet0" + str(x) + ".dwg"
            os.startfile(dir)
            if "acad.exe" in (i.name() for i in psutil.process_iter()):
                acad = Autocad()

                if running == True:
                    # Drawing 2 scripts
                    if x == 2:
                        if ul == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\02-noULScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\02-ULScript.scr \n')
                                    break
                                except:
                                    pass
                else:
                    return

                # Drawing 3 scripts
                if x == 3:
                    if running == True:
                        # Scale script
                        if scale == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\03-noScaleScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\03-scaleScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Auto film script
                        if auto == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\03-noAutoFilmScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\03-autoFilmScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return

                # Drawing 5 scripts
                if x == 5:
                    if running == True:
                        if split == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\05-noSplitFrameScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\05-splitFrameScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return

                # Drawing 7 scripts
                if x == 7:
                    if running == True:
                        # Cold Scripts
                        if cold == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\07-noColdOtherScript.scr \n')
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
                                    acad.doc.SendCommand('Script ' + scriptDir + '\\07-coldOtherScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Synergy Scripts
                        if syn == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\07-syn2Script.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\07-otherSynScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return

                if running == True:
                    # Drawings 02 - 09 edit script
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                            break
                        except:
                            pass
                else:
                    return
            # autocad not running error
            else:
                return
    else:
        return

    if running == True:
        # Drawings 10 - 27
        for x in range(10, 28):
            dir = folder + "\\" + projNum + "_Sheet" + str(x) + ".dwg"
            os.startfile(dir)
            if "acad.exe" in (i.name() for i in psutil.process_iter()):
                acad = Autocad()

                # Drawing 15 scripts
                if x == 15:
                    if running == True:
                        # Synergy script
                        if syn == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-syn2Script.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-otherSynScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Photo eye script
                        if photo == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-noPhotoEyeScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-photoEyeScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Carriage door script
                        if door == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-noCarriageDoorScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-carriageDoorScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        if not (syn == 0):
                            # HAR script
                            if har == 0:
                                timer = time.perf_counter()
                                while True:
                                    endTimer = time.perf_counter()
                                    if (endTimer - timer) >= 60:
                                        return
                                    try:
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-noHARScript.scr \n')
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
                                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\15-HARScript.scr \n')
                                        break
                                    except:
                                        pass
                    else:
                        return

                # Drawing 16 scripts
                if x == 16:
                    if running == True:
                        # Synergy script
                        if syn == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\16-syn2Script.scr \n')
                                    break
                                except:
                                    pass
                        elif syn == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\16-syn2_5Script.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\16-syn3Script.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Cold package script
                        if cold == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\16-noColdScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\16-coldScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return

                # Drawing 17 scripts
                if x == 17:
                    if running == True:
                        # Auto film scripts
                        if auto == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\17-noAutoFilmScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\17-autoFilmScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                
                # Drawing 25 scripts
                if x == 25:
                    if running == True:
                        # Synergy scripts
                        if syn == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-syn2Script.scr \n')
                                    break
                                except:
                                    pass
                        elif syn == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-syn2_5Script.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-syn3Script.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Carriage door scripts
                        if door == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noCarriageDoorScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-carriageDoorScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # HAR scripts
                        if har == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noHARScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-HARScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                    if running == True:
                        # Cold and auto film scripts
                        if auto == 1 and cold == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-coldAutoFilmScript \n')
                                    break
                                except:
                                    pass
                        elif auto == 0 and cold == 1:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-coldNoAutoScript.scr \n')
                                    break
                                except:
                                    pass
                        elif auto == 1 and cold == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noColdAutoScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\25-noColdNoAutoFilmScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return

                # Drawing 26 scripts
                if x == 26:
                    if running == True:
                        # Cold scripts
                        if cold == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\26-noColdScript.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\26-coldScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                
                # Drawing 27 scripts
                if x == 27:
                    if running == True:
                        # Synergy scripts
                        if syn == 0:
                            timer = time.perf_counter()
                            while True:
                                endTimer = time.perf_counter()
                                if (endTimer - timer) >= 60:
                                    return
                                try:
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\27-syn2Script.scr \n')
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
                                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\27-otherSynScript.scr \n')
                                    break
                                except:
                                    pass
                    else:
                        return
                
                if running == True:
                    # Drawings 10 - 27 edit script
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                            break
                        except:
                            pass
                else:
                    return

            # autocad not running error
            else:
                return
    else:
        return


    # Drawing 28
    dir = folder + "\\" + projNum + "_Sheet28.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
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
                if (endTimer - timer) >= 60:
                    return
                try:
                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                    break
                except:
                    pass
        # autocad not running error
        else:
            return
    else:
        return


    # Drawing 29
    dir = folder + "\\" + projNum + "_Sheet29.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
            
            if running == True:
                # Cold package scripts
                if cold == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noColdScript.scr \n')
                            break
                        except:
                            pass
                    # Cold + Synergy scripts
                    if syn == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noColdSyn2Script.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noColdOtherScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-coldScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Scale package script
                if scale == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noScaleScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-scaleScript.scr \n')
                            break
                        except:
                            pass
            else:
                return
            
            if running == True:
                # Advanced photo eye scripts
                if photo == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noPhotoEyeScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-photoEyeScript.scr \n')
                            break
                        except:
                            pass
            else:
                return
            
            if running == True:
                # Auto film scripts
                if auto == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noAutoFilmScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-autoFilmScript.scr \n')
                            break
                        except:
                            pass
                    # Auto + Cold scripts
                    if cold == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-autoNoColdScript.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-autoColdScript.scr \n')
                                break
                            except:
                                pass
            else:
                return

            if running == True:
                # UL script
                if ul == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULScript.scr \n')
                            break
                        except:
                            pass
                    # UL + Scale script
                    if scale == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULNoScaleScript.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULScaleScript.scr \n')
                                break
                            except:
                                pass
                    # UL + Auto script
                    if auto == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULNoAutoScript.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-noULAutoScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-ULScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Synergy scripts
                if syn == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-syn2Script.scr \n')
                            break
                        except:
                            pass
                    # HAR + Carriage door scripts with synergy 2
                    if har == 0 and door == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-0Script.scr \n')
                                break
                            except:
                                pass
                    elif (har == 1 and door == 0) or (har == 0 and door == 1):
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-1Script.scr \n')
                                break
                            except:
                                pass
                    else:
                        pass
                else:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-otherSynScript.scr \n')
                            break
                        except:
                            pass
                    # HAR + Carriage door scripts with synergy 2.5 and 3
                    if har == 0 and door == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-1Script.scr \n')
                                break
                            except:
                                pass
                    elif (har == 1 and door == 0) or (har == 0 and door == 1):
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\29-51-2Script.scr \n')
                                break
                            except:
                                pass
                    else:
                        pass
            else:
                return

            if running == True:
                # Edit script
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
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return

        # autocad not running error
        else:
            return
    else:
        return
        

    # Drawing 31
    dir = folder + "\\" + projNum + "_Sheet31.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            if running == True:
                # Synergy script
                if syn == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\31-syn2Script.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\31-otherSynScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Cold script
                if cold == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\31-noColdScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\31-coldScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Edit script
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return

        # autocad not running error
        else:
            return
    else:
        return
        

    # Drawing 32
    dir = folder + "\\" + projNum + "_Sheet32.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            if running == True:
                # Synergy scripts
                if syn == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-syn2Script.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-otherSynScript.scr \n')
                            break
                        except:
                            pass
            else:
                return
            
            if running == True:
                # Cold package scripts
                if cold == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-noColdScript.scr \n')
                            break
                        except:
                            pass
                    # Synergy + Cold scripts
                    if syn == 0:
                        timer = time.perf_counter()
                        while True:
                            endTimer = time.perf_counter()
                            if (endTimer - timer) >= 60:
                                return
                            try:
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-syn2NoColdScript.scr \n')
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
                                acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-synONoColdScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-coldScript.scr \n')
                            break
                        except:
                            pass
            else:
                return
            
            if running == True:
                # Split frame scripts
                if split == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-noSplitFrameScript.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\32-splitFrameScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:      
                # Edit script
                timer = time.perf_counter()
                while True:
                    endTimer = time.perf_counter()
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return
                
        # autocad not running error
        else:
            return
    else:
        return

    # Drawing 30
    dir = folder + "\\" + projNum + "_Sheet30.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()

            if running ==  True:
                # Synergy script
                if syn == 0:
                    timer = time.perf_counter()
                    while True:
                        endTimer = time.perf_counter()
                        if (endTimer - timer) >= 60:
                            return
                        try:
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\30-syn2Script.scr \n')
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
                            acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\30-otherSynScript.scr \n')
                            break
                        except:
                            pass
            else:
                return

            if running == True:
                # Edit script
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
                    if (endTimer - timer) >= 60:
                        return
                    try:
                        acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\tempScript.scr \n')
                        break
                    except:
                        pass
            else:
                return
        
        # autocad not running error
        else:
            return
    else:
        return


    # End scripts 
    dir = folder + "\\" + projNum + "_Sheet00.dwg"
    if running == True:
        os.startfile(dir)
        if "acad.exe" in (i.name() for i in psutil.process_iter()):
            acad = Autocad()
            timer = time.perf_counter()
            while True:
                endTimer = time.perf_counter()
                if (endTimer - timer) >= 60:
                    return
                try:
                    acad.doc.SendCommand('SCRIPT ' + scriptDir + '\\endScript.scr \n')
                    break
                except:
                    pass
        # autocad not running error
        else:
            return
    else:
        return

    # message sent to user telling them that the drawings have been edited for jobNum
    tk.messagebox.showinfo(title="Finished Editing Drawings", message="Job " + projNum + " has been created in '" + projNum + " Drawings' folder")
