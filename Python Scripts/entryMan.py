import tkinter as tk
import directoryFinder
from tkinter import *

########################################
# Disables directory entry box         #
########################################
def disableDEntry(jobEntry, dirEntry):
    dir = directoryFinder.findDir(jobEntry.get()) + ""
    dirEntry.config(state="normal")
    dirEntry.delete(0, tk.END)
    if dir == "":
        jobEntry.delete(0, tk.END)
        dirEntry.insert(0, "Waiting...")
        dirEntry.config(state="disabled")
        tk.messagebox.showerror(title="Invalid Input", message="invalid job number")
        return
    elif dir == "Short" or dir == "Empty":
        dirEntry.insert(0, "Waiting...")
    elif dir == "Long":
        dirEntry.insert(0, "Too long...")
    else:
        dirEntry.insert(0, dir)
    dirEntry.config(state="disabled")

#######################################
# Enables directory entry box         #
#######################################
def enableDEntry(jobEntry, dirEntry):
    dir = directoryFinder.findDir(jobEntry.get()) + ""
    dirEntry.config(state="normal")
    dirEntry.delete(0, tk.END)
    if dir == "":
        pass
    elif dir == "Short" or dir == "Empty":
        pass
    elif dir == "Long":
        pass
    else:
        dirEntry.insert(0, dir)

############################################
# Disables access directory entry box      #
############################################
def disableADEntry(accessJobEntry, accessDirEntry):
    tempJN = accessJobEntry.get()
    dir = directoryFinder.findDir(tempJN) + ""
    accessDirEntry.config(state="normal")
    accessDirEntry.delete(0, tk.END)
    if dir == "":
        accessJobEntry.delete(0, tk.END)
        accessDirEntry.insert(0, "Waiting...")
        accessDirEntry.config(state="disabled")
        tk.messagebox.showerror(title="Invalid Input", message="invalid job number")
        return
    elif dir == "Short" or dir == "Empty":
        accessDirEntry.insert(0, "Waiting...")
    elif dir == "Long":
        accessDirEntry.insert(0, "Too long...")
    else:
        dir = dir.replace(tempJN + " Drawings", tempJN + " Electrical BOM")
        accessDirEntry.insert(0, dir)
    accessDirEntry.config(state="disabled")

############################################
# Enables access directory entry box       #
############################################
def enableADEntry(accessJobEntry, accessDirEntry):
    tempJN = accessJobEntry.get()
    dir = directoryFinder.findDir(tempJN) + ""
    accessDirEntry.config(state="normal")
    accessDirEntry.delete(0, tk.END)
    if dir == "":
        pass
    elif dir == "Short" or dir == "Empty" or dir == "Long":
        pass
    else:
        dir = dir.replace(tempJN + " Drawings", tempJN + " Electrical BOM")
        accessDirEntry.insert(0, dir)

##################################################
# Disables export excel directory entry box      #
##################################################
def disableExDEntry(exJobNumEntry, exDirEntry):
    tempJN = exJobNumEntry.get()
    dir = directoryFinder.findDir(tempJN) + ""
    exDirEntry.config(state="normal")
    exDirEntry.delete(0, tk.END)
    if dir == "":
        exJobNumEntry.delete(0, tk.END)
        exDirEntry.insert(0, "Waiting...")
        exDirEntry.config(state="disabled")
        tk.messagebox.showerror(title="Invalid Input", message="invalid job number")
        return
    elif dir == "Short" or dir == "Empty":
        exDirEntry.insert(0, "Waiting...")
    elif dir == "Long":
        exDirEntry.insert(0, "Too long...")
    else:
        dir = dir.replace(tempJN + " Drawings", tempJN + " Electrical BOM")
        exDirEntry.insert(0, dir)
    exDirEntry.config(state="disabled")

##################################################
# Enables export excel directory entry box       #
##################################################
def enableExDEntry(exJobNumEntry, exDirEntry):
    tempJN = exJobNumEntry.get()
    dir = directoryFinder.findDir(tempJN) + ""
    exDirEntry.config(state="normal")
    exDirEntry.delete(0, tk.END)
    if dir == "":
        pass
    elif dir == "Short" or dir == "Empty" or dir == "Long":
        pass
    else:
        dir = dir.replace(tempJN + " Drawings", tempJN + " Electrical BOM")
        exDirEntry.insert(0, dir)

############################################
# Disables pages entry box                 #
############################################
def disablePEntry(pagesEntry):
    pagesEntry.config(state="disable")

############################################
# Enables pages entry box                  #
############################################
def enablePEntry(pagesEntry):
    pagesEntry.config(state="normal")

########################################################
# Changes access dir based on job number entry         #
########################################################
def accDirEntryUpdate(event, accessJobVar, accessJobEntry, accessDirEntry):
    if accessJobVar.get() == 0:
        tempJN = accessJobEntry.get()
        dir = directoryFinder.findDir(tempJN) + ""
        accessDirEntry.config(state="normal")
        accessDirEntry.delete(0, tk.END)
        if dir == "":
            accessDirEntry.insert(0, "Invalid...")
        elif dir == "Short" or dir == "Empty":
            accessDirEntry.insert(0, "Waiting...")
        elif dir == "Long":
            accessDirEntry.insert(0, "Too long...")
        else:
            dir = dir.replace(tempJN + " Drawings", tempJN + " Electrical BOM")
            accessDirEntry.insert(0, dir)
        accessDirEntry.config(state="disabled")
    else: pass

##############################################################
# Changes export excel dir based on job number entry         #
##############################################################
def exDirEntryUpdate(event, exportJobVar, exJobNumEntry, exDirEntry):
    if exportJobVar.get() == 0:
        tempJN = exJobNumEntry.get()
        dir = directoryFinder.findDir(tempJN) + ""
        exDirEntry.config(state="normal")
        exDirEntry.delete(0, tk.END)
        if dir == "":
            exDirEntry.insert(0, "Invalid...")
        elif dir == "Short" or dir == "Empty":
            exDirEntry.insert(0, "Waiting...")
        elif dir == "Long":
            exDirEntry.insert(0, "Too long...")
        else:
            dir = dir.replace(tempJN + " Drawings", tempJN + " Electrical BOM")
            exDirEntry.insert(0, dir)
        exDirEntry.config(state="disabled")
    else: pass

####################################################
# Changes dir based on job number entry            #
####################################################
def dirEntryUpdate(event, jobVar, jobEntry, dirEntry):
    if jobVar.get() == 0:
        dir = directoryFinder.findDir(jobEntry.get()) + ""
        dirEntry.config(state="normal")
        dirEntry.delete(0, tk.END)
        if dir == "":
            dirEntry.insert(0, "Invalid...")
        elif dir == "Short" or dir == "Empty":
            dirEntry.insert(0, "Waiting...")
        elif dir == "Long":
            dirEntry.insert(0, "Too long...")
        else:
            dirEntry.insert(0, dir)
        dirEntry.config(state="disabled")
    else: pass