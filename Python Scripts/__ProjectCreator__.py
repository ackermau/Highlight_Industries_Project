import tkinter as tk
import entryMan
import callBack
from win32 import win32print
from tkinter import *
from tkinter import messagebox
from tkinter import ttk

#############################
# collapsible frame code    #
#############################
class ToggledFrame(tk.Frame):

    def __init__(self, parent, text="", *args, **options):
        tk.Frame.__init__(self, parent, *args, **options)

        self.show = tk.IntVar()
        self.show.set(0)

        self.title_frame = ttk.Frame(master=self)
        self.title_frame.pack(fill="x", expand=1)

        ttk.Label(master=self.title_frame, text=text, style="BW.TLabel").pack(side="left", fill="x", expand=1)

        self.toggle_button = ttk.Checkbutton(master=self.title_frame, width=5, text='    +', command=self.toggle, variable=self.show, style='BW.TLabel')
        self.toggle_button.pack(side="left", )

        self.sub_frame = tk.Frame(master=self, relief="sunken", borderwidth=1, background=machineFrameColor)

    def toggle(self):
        if bool(self.show.get()):
            self.sub_frame.pack(fill="x", expand=1)
            self.toggle_button.configure(text='    -')
            toggle = '-'
        else:
            self.sub_frame.forget()
            self.toggle_button.configure(text='    +')
            toggle = '+'
        resizeWindowMan(toggle=toggle)

###############################################
# Changes machine specs to synergy machines   #
###############################################
def synConfig():
    if macTypeVar.get() == 0:
        syn2RadioB["text"] = "Synergy 2   "
        syn2RadioB["variable"] = synVar
        syn2_5RadioB.pack()
        syn3RadioB.pack()

        hiProfileCB.grid(column=0, row=0, sticky=W, padx=5, pady=5)
        lowProfileCB.grid(column=0, row=1, sticky=W, padx=5, pady=5)
        scaleCB.grid(column=0, row=2, sticky=W, padx=5, pady=5)
        coldCB.grid(column=0, row=3, sticky=W, padx=5, pady=5)
        splitCB.grid(column=0, row=4, sticky=W, padx=5, pady=5)
        AMP20CB.grid(column=0, row=5, sticky=W, padx=5, pady=5)
        autoCB.grid(column=1, row=0, sticky=W, padx=5, pady=5)
        ulCB.grid(column=1, row=1, sticky=W, padx=5, pady=5)
        doorCB.grid(column=1, row=2, sticky=W, padx=5, pady=5)
        harCB.grid(column=1, row=3, sticky=W, padx=5, pady=5)
        add9CB.grid(column=1, row=4, sticky=W, padx=5, pady=5)
        add10CB.grid(column=1, row=5, sticky=W, padx=5, pady=5)
        
    elif macTypeVar.get() == 1:
        syn2RadioB["text"] = "Synergy 5   "
        syn2RadioB["variable"] = autoSynVar

        autoAdd1CB.grid(column=0, row=0, sticky=W, padx=5, pady=5)
        autoAdd2CB.grid(column=0, row=1, sticky=W, padx=5, pady=5)
        autoAdd3CB.grid(column=0, row=2, sticky=W, padx=5, pady=5)
        autoAdd4CB.grid(column=0, row=3, sticky=W, padx=5, pady=5)
        autoAdd5CB.grid(column=0, row=4, sticky=W, padx=5, pady=5)
        autoAdd6CB.grid(column=0, row=5, sticky=W, padx=5, pady=5)
        autoAdd7CB.grid(column=1, row=0, sticky=W, padx=5, pady=5)
        autoAdd8CB.grid(column=1, row=1, sticky=W, padx=5, pady=5)
        autoAdd9CB.grid(column=1, row=2, sticky=W, padx=5, pady=5)
        autoAdd10CB.grid(column=1, row=3, sticky=W, padx=5, pady=5)
        autoAdd11CB.grid(column=1, row=4, sticky=W, padx=5, pady=5)
        autoAdd12CB.grid(column=1, row=5, sticky=W, padx=5, pady=5)

###############################################
# Changes machine specs to freedom machines   #
###############################################
def freeConfig():
    if macTypeVar.get() == 0:
        syn2RadioB["text"] = "Freedom Semi"
        syn2RadioB["variable"] = freeVar
        syn2_5RadioB.pack_forget()
        syn3RadioB.pack_forget()

        hiProfileCB.grid_forget()
        lowProfileCB.grid_forget()
        scaleCB.grid_forget()
        coldCB.grid_forget()
        splitCB.grid_forget()
        AMP20CB.grid_forget()
        autoCB.grid_forget()
        ulCB.grid_forget()
        doorCB.grid_forget()
        harCB.grid_forget()
        add9CB.grid_forget()
        add10CB.grid_forget()

    elif macTypeVar.get() == 1:
        syn2RadioB["text"] = "Freedom 6500"
        syn2RadioB["variable"] = autoFreeVar

        autoAdd1CB.grid_forget()
        autoAdd2CB.grid_forget()
        autoAdd3CB.grid_forget()
        autoAdd4CB.grid_forget()
        autoAdd5CB.grid_forget()
        autoAdd6CB.grid_forget()
        autoAdd7CB.grid_forget()
        autoAdd8CB.grid_forget()
        autoAdd9CB.grid_forget()
        autoAdd10CB.grid_forget()
        autoAdd11CB.grid_forget()
        autoAdd12CB.grid_forget()

#########################################################
# Changes main tab to automatic machine configuration   #
#########################################################
def autoConfig():
    # Button manipulation
    autoB.config(relief=SUNKEN, bg=machineFrameColor, fg=textColor)
    semiAutoB.config(relief=RAISED, bg=buttonColor, fg=machineFrameColor)

    # Label manipulation
    if macTypeVar.get() == 0:
        # Setting main tab for automatic machines
        synRadioB["text"] = "Synergy - Automatic  "
        synRadioB["variable"] = macVarA
        freeRadioB["text"] = "Freedom - Automatic"
        freeRadioB["variable"] = macVarA
        if macVarA.get() == 0:
            syn2RadioB["text"] = "Synergy 5   "
            syn2RadioB["variable"] = autoSynVar
        elif macVarA.get() == 1:
            syn2RadioB["text"] = "Freedom 6500"
            syn2RadioB["variable"] = autoFreeVar
        syn2_5RadioB.pack_forget()
        syn3RadioB.pack_forget()

        hiProfileCB.grid_forget()
        lowProfileCB.grid_forget()
        scaleCB.grid_forget()
        coldCB.grid_forget()
        splitCB.grid_forget()
        AMP20CB.grid_forget()
        autoCB.grid_forget()
        ulCB.grid_forget()
        doorCB.grid_forget()
        harCB.grid_forget()
        add9CB.grid_forget()
        add10CB.grid_forget()

        autoAdd1CB.grid(column=0, row=0, sticky=W, padx=5, pady=5)
        autoAdd2CB.grid(column=0, row=1, sticky=W, padx=5, pady=5)
        autoAdd3CB.grid(column=0, row=2, sticky=W, padx=5, pady=5)
        autoAdd4CB.grid(column=0, row=3, sticky=W, padx=5, pady=5)
        autoAdd5CB.grid(column=0, row=4, sticky=W, padx=5, pady=5)
        autoAdd6CB.grid(column=0, row=5, sticky=W, padx=5, pady=5)
        autoAdd7CB.grid(column=1, row=0, sticky=W, padx=5, pady=5)
        autoAdd8CB.grid(column=1, row=1, sticky=W, padx=5, pady=5)
        autoAdd9CB.grid(column=1, row=2, sticky=W, padx=5, pady=5)
        autoAdd10CB.grid(column=1, row=3, sticky=W, padx=5, pady=5)
        autoAdd11CB.grid(column=1, row=4, sticky=W, padx=5, pady=5)
        autoAdd12CB.grid(column=1, row=5, sticky=W, padx=5, pady=5)

        macTypeVar.set(1)
    else:
        pass

#########################################################
# Changes main tab to semi auto machine configuration   #
#########################################################
def semiConfig():
    # Button manipulation
    autoB.config(relief=RAISED, bg=buttonColor, fg=machineFrameColor)
    semiAutoB.config(relief=SUNKEN, bg=machineFrameColor, fg=textColor)

    # Lable manipulation
    if macTypeVar.get() == 1:
        # Setting main tab for semi auto machines
        synRadioB["text"] = "Synergy - Semi Auto "
        synRadioB["variable"] = macVarS
        freeRadioB["text"] = "Freedom - Semi Auto"
        freeRadioB["variable"] = macVarS
        if macVarS.get() == 0:
            syn2RadioB["text"] = "Synergy 2   "
            syn2RadioB["variable"] = synVar
            syn2_5RadioB.pack()
            syn3RadioB.pack()
        elif macVarS.get() == 1:
            syn2RadioB["text"] = "Freedom Semi"
            syn2RadioB["variable"] = freeVar
        
        autoAdd1CB.grid_forget()
        autoAdd2CB.grid_forget()
        autoAdd3CB.grid_forget()
        autoAdd4CB.grid_forget()
        autoAdd5CB.grid_forget()
        autoAdd6CB.grid_forget()
        autoAdd7CB.grid_forget()
        autoAdd8CB.grid_forget()
        autoAdd9CB.grid_forget()
        autoAdd10CB.grid_forget()
        autoAdd11CB.grid_forget()
        autoAdd12CB.grid_forget()

        hiProfileCB.grid(column=0, row=0, sticky=W, padx=5, pady=5)
        lowProfileCB.grid(column=0, row=1, sticky=W, padx=5, pady=5)
        scaleCB.grid(column=0, row=2, sticky=W, padx=5, pady=5)
        coldCB.grid(column=0, row=3, sticky=W, padx=5, pady=5)
        splitCB.grid(column=0, row=4, sticky=W, padx=5, pady=5)
        AMP20CB.grid(column=0, row=5, sticky=W, padx=5, pady=5)
        autoCB.grid(column=1, row=0, sticky=W, padx=5, pady=5)
        ulCB.grid(column=1, row=1, sticky=W, padx=5, pady=5)
        doorCB.grid(column=1, row=2, sticky=W, padx=5, pady=5)
        harCB.grid(column=1, row=3, sticky=W, padx=5, pady=5)
        add9CB.grid(column=1, row=4, sticky=W, padx=5, pady=5)
        add10CB.grid(column=1, row=5, sticky=W, padx=5, pady=5)

        macTypeVar.set(0)
    else:
        pass

############################################
# Resizes window when tab is changed       #
############################################
def resizeWindow(*args):
    selectedTab = tabControl.select()
    tabText = tabControl.tab(selectedTab, "text")

    if tabText == "Main":
        window.geometry("")
        custEntry.focus()
    elif tabText == "Data":
        window.geometry("358x382")
        exJobNumEntry.focus()
    elif tabText == "Print":
        window.geometry("358x368")
        jobEntry.focus()
    elif tabText == "About":
        window.geometry("358x112")
    else:
        window.geometry("")

##############################################
# Resizes window when tab is changed         #
##############################################
def resizeWindowMan(toggle):
    selectedTab = tabControl.select()
    tabText = tabControl.tab(selectedTab, "text")

    if tabText == "Main":
        window.geometry("")
        custEntry.focus()
    elif tabText == "Data":
        window.geometry("358x382")
        exJobNumEntry.focus()
    elif tabText == "Print":
        window.geometry("358x368")
        jobEntry.focus()
    elif tabText == "About":
        window.geometry("358x112")
    else:
        window.geometry("")

####################################
# main tkinter method code         #
####################################
window = tk.Tk()
window.title("Standard Machine Creator")
window.geometry("")
window.resizable(False, False)
numRan = 0
try:
    window.iconbitmap('./Python Scripts/h_lg.ico')
except:
    tk.messagebox.showwarning(title="Failed Loading Icon", message="Icon file was corrupted or lost")
machineFrameColor = "#393E46"
textColor = "#EEEEEE"
buttonColor = "#00ADB5"
miscColor = "#AAD8D3"
textFont = ('Space Grotesk', 11, 'bold')
miniFont = ('Space Grotesk', 11)
tabControl = ttk.Notebook(window)

########################
# Tab control frame    #
########################
tabStyle = ttk.Style()
tabStyle.theme_use('default')
tabStyle.configure('TNotebook.Tab', background=machineFrameColor)
tabStyle.configure('TNotebook.Tab', foreground=textColor)
tabStyle.configure('TNotebook.Tab', font=miniFont)
tabStyle.map("TNotebook.Tab", background=[("selected", buttonColor)])
tabStyle.map("TNotebook.Tab", foreground=[("selected", "black")])
style = ttk.Style()
style.configure("BW.TLabel", foreground=miscColor, background=machineFrameColor, font=textFont)
mainTab = ttk.Frame(tabControl, style="BW.TLabel")
exportTab = ttk.Frame(tabControl, style="BW.TLabel")
printTab = ttk.Frame(tabControl, style="BW.TLabel")
infoTab = ttk.Frame(tabControl, style="BW.TLabel")
tabControl.add(mainTab, text='Main')
tabControl.add(exportTab, text='Data')
tabControl.add(printTab, text='Print')
tabControl.add(infoTab, text='About')
tabControl.bind('<<NotebookTabChanged>>', resizeWindow)
tabControl.pack(expand=1, fill="both")

#############################
# Machine ID tab            #
#############################
machineIDFrame = tk.Frame(master=mainTab, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)

# Machine button frame
macButtonFrame = tk.Frame(master=machineIDFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
macButtonFrame.columnconfigure(0, weight=1)
macButtonFrame.columnconfigure(1, weight=1)

# Information frame
infoFrame = tk.Frame(master=machineIDFrame, bg=machineFrameColor)
infoFrame.config(height=50, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
infoFrame.columnconfigure(0, weight=1)
infoFrame.columnconfigure(1, weight=3)

# Customer Label/box
custLabel = tk.Label(master=infoFrame, text="Customer :", font=textFont, bg=machineFrameColor, fg=textColor)
custLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
custEntry = tk.Entry(master=infoFrame, font=textFont, bg=miscColor, fg=machineFrameColor)
custEntry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

# Distributor Label/box
distrLabel = tk.Label(master=infoFrame, text="Distributor :", font=textFont, bg=machineFrameColor, fg=textColor)
distrLabel.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
distrEntry = tk.Entry(master=infoFrame, font=textFont, bg=miscColor, fg=machineFrameColor)
distrEntry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

# Project Label/box
projNumLabel = tk.Label(master=infoFrame, text="Project # :", font=textFont, bg=machineFrameColor, fg=textColor)
projNumLabel.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
projNumEntry = tk.Entry(master=infoFrame, font=textFont, bg=miscColor, fg=machineFrameColor)
projNumEntry.grid(column=1, row=2, sticky=tk.E, padx=5, pady=5)

# Manufacturing Year Label/box
manYearLabel = tk.Label(master=infoFrame, text="Manufactured Year : ", font=textFont, bg=machineFrameColor, fg=textColor)
manYearLabel.grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)
manYearEntry = tk.Entry(master=infoFrame, font=textFont, bg=miscColor, fg=machineFrameColor)
manYearEntry.grid(column=1, row=3, sticky=tk.E, padx=5, pady=5)

#
# User Information labels, and text entries
#
# Engineer label/box
enginLabel = tk.Label(master=infoFrame, text="Engineer (Initials) :", font=textFont, bg=machineFrameColor, fg=textColor)
enginLabel.grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
enginEntry = tk.Entry(master=infoFrame, font=textFont, bg=miscColor, fg=machineFrameColor)
enginEntry.grid(column=1, row=4, sticky=tk.E, padx=5, pady=5)

# Date label/box
dateLabel = tk.Label(master=infoFrame, text="Date (mm/dd/yyyy) :", font=textFont, bg=machineFrameColor, fg=textColor)
dateLabel.grid(column=0, row=5, sticky=tk.W, padx=5, pady=5)
dateEntry = tk.Entry(master=infoFrame, font=textFont, bg=miscColor, fg=machineFrameColor)
dateEntry.grid(column=1, row=5, sticky=tk.E, padx=5, pady=5)

#
# Machine Type radio buttons and frame
#
macTypeVar = IntVar()
# macTypeVar is 0 = Semi Auto if macTypeVar is 1 = Automatic
macTypeVar.set(0)
macVarS = IntVar()
macVarS.set(0)
macVarA = IntVar()
macVarA.set(0)

# Machine type frame
machTypeFrame = tk.Frame(master=machineIDFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)

machStyleFrame = tk.Frame(master=machTypeFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)

semiAutoB = Button(master=machStyleFrame, text="Semi Auto Machines", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, relief=SUNKEN, bg=machineFrameColor, fg=textColor, command=semiConfig)
semiAutoB.grid(column=0, row=0, padx=5, pady=5)

autoB = Button(master=machStyleFrame, text="Automatic Machines", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command=autoConfig)
autoB.grid(column=0, row=1, padx=5, pady=5)

synRadioB = Radiobutton(master=machStyleFrame, text="Synergy - Semi Auto ", font=textFont, variable=macVarS, value=0, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor, command=synConfig)
synRadioB.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)

freeRadioB = Radiobutton(master=machStyleFrame, text="Freedom - Semi Auto", font=textFont, variable=macVarS, value=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor, command=freeConfig)
freeRadioB.grid(column=1, row=1, sticky=tk.W, padx=5, pady=5)

#
# Synergy machine type frame
#
synFrame = ToggledFrame(parent=machTypeFrame, text="Machine Type", relief="raised", borderwidth=1)
synFrame.config(bg=machineFrameColor)

synVar = IntVar()
autoSynVar = IntVar()
freeVar = IntVar()
autoFreeVar = IntVar()

syn2RadioB = Radiobutton(master=synFrame.sub_frame, text="Synergy 2   ", font=miniFont, variable=synVar, value=0, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
syn2RadioB.pack()

syn2_5RadioB = Radiobutton(master=synFrame.sub_frame, text="Synergy 2.5", font=miniFont, variable=synVar, value=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
syn2_5RadioB.pack()

syn3RadioB = Radiobutton(master=synFrame.sub_frame, text="Synergy 3   ", font=miniFont, variable=synVar, value=2, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
syn3RadioB.pack()

syn4RadioB = Radiobutton(master=synFrame.sub_frame, text="Synergy 4   ", font=miniFont, variable=synVar, value=3, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
syn4RadioB.pack()

#
# Optional additions frame
#
opAddFrame = ToggledFrame(parent=machTypeFrame, text="Optional Additions", relief="raised", borderwidth=1)
opAddFrame.config(bg=machineFrameColor)
opAddFrame.columnconfigure(0, weight=1)
opAddFrame.columnconfigure(1, weight=1)

# Hi profile check button and variable
profileVar = IntVar()
hiProfileCB = Radiobutton(master=opAddFrame.sub_frame, text="High Profile", font=miniFont, variable=profileVar, value=0, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
hiProfileCB.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

# Low profile check button and variable
lowProfileCB = Radiobutton(master=opAddFrame.sub_frame, text="Low Profile", font=miniFont, variable=profileVar, value=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
lowProfileCB.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)

# Scale package check button and variable
scaleVar = IntVar()
scaleCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Scale package", font=miniFont, variable=scaleVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
scaleCB.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)

# Cold package check button and variable
coldVar = IntVar()
coldCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Cold package", font=miniFont, variable=coldVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
coldCB.grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)

# Split frame check button and variable
splitVar = IntVar()
splitCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Split frame", font=miniFont, variable=splitVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
splitCB.grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)

# Additional option 9
AMP20V = IntVar()
AMP20CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="20 AMP", font=miniFont, variable=AMP20V, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
AMP20CB.grid(column=0, row=5, sticky=tk.W, padx=5, pady=5)

# Auto film cut check button and variable
autoVar = IntVar()
autoCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Auto Film Cut", font=miniFont, variable=autoVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
autoCB.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)

# UL check button and variable
ulVar = IntVar()
ulCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="UL Option", font=miniFont, variable=ulVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
ulCB.grid(column=1, row=1, sticky=tk.W, padx=5, pady=5)

# Carriage door checkbutton and variable
doorVar = IntVar()
doorCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Carriage Door SW", font=miniFont, variable=doorVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
doorCB.grid(column=1, row=2, sticky=tk.W, padx=5, pady=5)

# HAR check button and variable
harVar = IntVar()
harCB = tk.Checkbutton(master=opAddFrame.sub_frame, text="HAR Option", font=miniFont, variable=harVar, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
harCB.grid(column=1, row=3, sticky=tk.W, padx=5, pady=5)

# Additional option 9
add9V = IntVar()
add9CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 9", font=miniFont, variable=add9V, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
add9CB.grid(column=1, row=4, sticky=tk.W, padx=5, pady=5)

# Additional option 10
add10V = IntVar()
add10CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 10", font=miniFont, variable=add10V, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)
add10CB.grid(column=1, row=5, sticky=tk.W, padx=5, pady=5)

# Automatic machine options
autoAdd1Var = IntVar()
autoAdd1CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 1", font=miniFont, variable=autoAdd1Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd2Var = IntVar()
autoAdd2CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 2", font=miniFont, variable=autoAdd2Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd3Var = IntVar()
autoAdd3CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 3", font=miniFont, variable=autoAdd3Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd4Var = IntVar()
autoAdd4CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 4", font=miniFont, variable=autoAdd4Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd5Var = IntVar()
autoAdd5CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 5", font=miniFont, variable=autoAdd5Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd6Var = IntVar()
autoAdd6CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 6", font=miniFont, variable=autoAdd6Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd7Var = IntVar()
autoAdd7CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 7", font=miniFont, variable=autoAdd7Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd8Var = IntVar()
autoAdd8CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 8", font=miniFont, variable=autoAdd8Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd9Var = IntVar()
autoAdd9CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 9", font=miniFont, variable=autoAdd9Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd10Var = IntVar()
autoAdd10CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 10", font=miniFont, variable=autoAdd10Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd11Var = IntVar()
autoAdd11CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 11", font=miniFont, variable=autoAdd11Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

autoAdd12Var = IntVar()
autoAdd12CB = tk.Checkbutton(master=opAddFrame.sub_frame, text="Additional 12", font=miniFont, variable=autoAdd12Var, offvalue=0, onvalue=1, bg=machineFrameColor, fg=textColor, activebackground=buttonColor, activeforeground=textColor, selectcolor=machineFrameColor)

#
# Utility Specs labels, text entries, and frame
#
utilitySpecsFrame = ToggledFrame(parent=machTypeFrame, text="Utility Specs", relief="raised", borderwidth=1)
utilitySpecsFrame.config(bg=machineFrameColor)
utilitySpecsFrame.columnconfigure(0, weight=1)
utilitySpecsFrame.columnconfigure(1, weight=3)

# Phase Label/box, default = 60 HERTZ
phaseLabel = tk.Label(master=utilitySpecsFrame.sub_frame, text="3-Phase, +Earth :", font=miniFont, bg=machineFrameColor, fg=textColor)
phaseLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
phaseEntry = tk.Entry(master=utilitySpecsFrame.sub_frame, font=miniFont, bg=miscColor, fg=machineFrameColor)
phaseEntry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)
phaseEntry.insert(0, "60 HERTZ")

# Main Line Voltage Label/box, default = 120 VAC, 1 PHASE
mainLineVLabel = tk.Label(master=utilitySpecsFrame.sub_frame, text="Main Line Voltage :", font=miniFont, bg=machineFrameColor, fg=textColor)
mainLineVLabel.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
mainLineVEntry = tk.Entry(master=utilitySpecsFrame.sub_frame, font=miniFont, bg=miscColor, fg=machineFrameColor)
mainLineVEntry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)
mainLineVEntry.insert(0, "120 VAC, 1 PHASE")

# Control Voltage Label/box, default = 120 VAC, 24VDC
controlVLabel = tk.Label(master=utilitySpecsFrame.sub_frame, text="Control Voltage :", font=miniFont, bg=machineFrameColor, fg=textColor)
controlVLabel.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
controlVEntry = tk.Entry(master=utilitySpecsFrame.sub_frame, font=miniFont, bg=miscColor, fg=machineFrameColor)
controlVEntry.grid(column=1, row=2, sticky=tk.E, padx=5, pady=5)
controlVEntry.insert(0, "120 VAC, 24VDC")

# Total Motor Full Load Label/box, default = 2 HP
totMotorLabel = tk.Label(master=utilitySpecsFrame.sub_frame, text="Total Motor Full Load :", font=miniFont, bg=machineFrameColor, fg=textColor)
totMotorLabel.grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)
totMotorEntry = tk.Entry(master=utilitySpecsFrame.sub_frame, font=miniFont, bg=miscColor, fg=machineFrameColor)
totMotorEntry.grid(column=1, row=3, sticky=tk.E, padx=5, pady=5)
totMotorEntry.insert(0, "2 HP")

# Full Load Current Label/box, default = 15 AMP
fullLoadLabel = tk.Label(master=utilitySpecsFrame.sub_frame, text="Full Load Current :", font=miniFont, bg=machineFrameColor, fg=textColor)
fullLoadLabel.grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
fullLoadEntry = tk.Entry(master=utilitySpecsFrame.sub_frame, font=miniFont, bg=miscColor, fg=machineFrameColor)
fullLoadEntry.grid(column=1, row=4, sticky=tk.E, padx=5, pady=5)
if coldVar == 0:
    fullLoadEntry.insert(0, "15 AMP")
else:
    fullLoadEntry.insert(0, "20 AMP")

#############################
# Export tab                #
#############################
exFrame = tk.Frame(master=exportTab, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)

# Export Frame
exportFrame = tk.Frame(master=exFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
exportFrame.columnconfigure(0, weight=1)
exportFrame.columnconfigure(1, weight=3)

# Eport button frame
exButtonFrame = tk.Frame(master=exFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
exButtonFrame.columnconfigure(0, weight=1)
exButtonFrame.columnconfigure(1, weight=1)

# Export jobNum
exJobNumLabel = tk.Label(master=exportFrame, text="Job Number :", font=textFont, bg=machineFrameColor, fg=textColor)
exJobNumLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
exJobNumEntry = tk.Entry(master=exportFrame, font=textFont, bg=miscColor, fg=machineFrameColor)

# Export dir
exDirLabel = tk.Label(master=exportFrame, text="Directory :", fg=textColor, bg=machineFrameColor, font=textFont)
exDirLabel.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
exDirEntry = tk.Entry(master=exportFrame, fg=machineFrameColor, bg=miscColor, font=textFont)
exDirEntry.insert(0, "Waiting...")
exDirEntry.config(state="disabled")
exDirEntry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

# Export job radio buttons
exportJobVar = IntVar()

# Automatic finder radiobutton
exAutoRadBut = tk.Radiobutton(master=exportFrame, text="Automatic Finder", variable=exportJobVar, value=0, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.disableExDEntry(exJobNumEntry, exDirEntry))
exAutoRadBut.grid(column=0, row=2, sticky=tk.W, padx=0, pady=0)

# Manual directory radiobutton
exManRadBut = tk.Radiobutton(master=exportFrame, text="Manual Directory", variable=exportJobVar, value=1, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.enableExDEntry(exJobNumEntry, exDirEntry))
exManRadBut.grid(column=0, row=3, sticky=tk.W, padx=0, pady=0)

# Binding export excel job entry to export excel directory to key release event
exJobNumEntry.bind('<KeyRelease>', lambda event: entryMan.exDirEntryUpdate(event, exportJobVar, exJobNumEntry, exDirEntry))
exJobNumEntry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

##############################
# Access Tab                 #
##############################
accessFrame = tk.Frame(master=exportTab, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)

# Access button frame
accessButtonFrame = tk.Frame(master=accessFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
accessButtonFrame.columnconfigure(0, weight=1)
accessButtonFrame.columnconfigure(0, weight=1)

# Access job number frame
accessJobFrame = tk.Frame(master=accessFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor,highlightthickness=2)
accessJobFrame.columnconfigure(0, weight=1)
accessJobFrame.columnconfigure(1, weight=3)

# Access job number label/box
accessJobLabel = tk.Label(master=accessJobFrame, text="Job Number :", fg=textColor, bg=machineFrameColor, font=textFont)
accessJobLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
accessJobEntry = tk.Entry(master=accessJobFrame, fg=machineFrameColor, bg=miscColor, font=textFont)

# Directory label/box
accessDirLabel = tk.Label(master=accessJobFrame, text="Directory :", fg=textColor, bg=machineFrameColor, font=textFont)
accessDirLabel.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
accessDirEntry = tk.Entry(master=accessJobFrame, fg=machineFrameColor, bg=miscColor, font=textFont)
accessDirEntry.insert(0, "Waiting...")
accessDirEntry.config(state="disabled")
accessDirEntry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

# Access job radio buttons
accessJobVar = IntVar()

# Automatic finder radiobutton
accAutoRadBut = tk.Radiobutton(master=accessJobFrame, text="Automatic Finder", variable=accessJobVar, value=0, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.disableADEntry(accessJobEntry, accessDirEntry))
accAutoRadBut.grid(column=0, row=2, sticky=tk.W, padx=0, pady=0)

# Manual directory radiobutton
accManRadBut = tk.Radiobutton(master=accessJobFrame, text="Manual Directory", variable=accessJobVar, value=1, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.enableADEntry(accessJobEntry, accessDirEntry))
accManRadBut.grid(column=0, row=3, sticky=tk.W, padx=0, pady=0)

# Binding access job entry to access directory to key release event
accessJobEntry.bind('<KeyRelease>', lambda event: entryMan.accDirEntryUpdate(event, accessJobVar, accessJobEntry, accessDirEntry))
accessJobEntry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

##############################
# Print Tab                  #
##############################
printFrame = tk.Frame(master=printTab, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)

# Print button frame
printButtonFrame = tk.Frame(master=printFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
printButtonFrame.columnconfigure(0, weight=1)
printButtonFrame.columnconfigure(1, weight=1)

# Job frame
jobFrame = tk.Frame(master=printFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
jobFrame.columnconfigure(0, weight=1)
jobFrame.columnconfigure(1, weight=3)

# Job number Label/box
jobLabel = tk.Label(master=jobFrame, text="Job Number :", fg=textColor, bg=machineFrameColor, font=textFont)
jobLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
jobEntry = tk.Entry(master=jobFrame, fg=machineFrameColor, bg=miscColor, font=textFont)

# Directory label/box
dirLabel = tk.Label(master=jobFrame, text="Directory :", fg=textColor, bg=machineFrameColor, font=textFont)
dirLabel.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
dirEntry = tk.Entry(master=jobFrame, fg=machineFrameColor, bg=miscColor, font=textFont)
dirEntry.insert(0, "Waiting...")
dirEntry.config(state="disabled")
dirEntry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

# Job radio buttons
jobVar = IntVar()

jobRadioButton = tk.Radiobutton(master=jobFrame, text="Automatic Finder", variable=jobVar, value=0, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.disableDEntry(jobEntry, dirEntry))
jobRadioButton.grid(column=0, row=2, sticky=tk.N, padx=0, pady=0)

dirRadioButton = tk.Radiobutton(master=jobFrame, text="Manual Directory", variable=jobVar, value=1, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.enableDEntry(jobEntry, dirEntry))
dirRadioButton.grid(column=0, row=3, sticky=tk.N, padx=0, pady=0)

# Binding jobEntry and dirEntry to key release event
jobEntry.bind('<KeyRelease>', lambda event: entryMan.dirEntryUpdate(event, jobVar, jobEntry, dirEntry))
jobEntry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

# Pages Frame
pageFrame = tk.Frame(master=printFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
pageFrame.columnconfigure(0, weight=1)
pageFrame.columnconfigure(1, weight=3)

# Pages Label/box
pagesLabel = tk.Label(master=pageFrame, text="Pages :", fg=textColor, bg=machineFrameColor, font=textFont)
pagesLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
pagesEntry = tk.Entry(master=pageFrame, state="normal", fg=machineFrameColor, bg=miscColor, font=textFont)
pagesEntry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)
pagesEntry.insert(0, "0-2,4-7,15-16,25,27-32")

# Print radio buttons
paVar = IntVar()

psRadioButton = tk.Radiobutton(master=pageFrame, text="Print selected", variable=paVar, value=0, fg=textColor, bg=machineFrameColor, selectcolor=machineFrameColor, font=textFont, command= lambda: entryMan.enablePEntry(pagesEntry))
psRadioButton.grid(column=0, row=1, sticky=tk.W, padx=0, pady=0)

paRadioButton = tk.Radiobutton(master=pageFrame, text="Print all", variable=paVar, value=1, fg=textColor, bg=machineFrameColor, font=textFont, selectcolor=machineFrameColor, command= lambda: entryMan.disablePEntry(pagesEntry))
paRadioButton.grid(column=0, row=2, sticky=tk.W, padx=0, pady=0)

# Printer frame
printerFrame = tk.Frame(master=printFrame, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=2)
printerFrame.columnconfigure(0, weight=1)

# Printer Label
printerLabel = tk.Label(master=printerFrame, text="Select Printer :", fg=textColor, bg=machineFrameColor, font=textFont)
printerLabel.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

# Printer combobox
printerVar = StringVar()
printerComboBox = ttk.Combobox(master=printerFrame, width=37, textvariable=printerVar, font=textFont)
printerComboBox.config(foreground=machineFrameColor)
printerList = []
printers = list(win32print.EnumPrinters(4))
for i in printers:
    printerList.append(i[2])
printerComboBox['values'] = printerList
for i in printerComboBox['values']:
    if i == '\\\\server\\EE - Sharp MX-3501N':
        printerComboBox.set(i)
    else: pass
if not (printerComboBox.current() == '\\\\server\\EE - Sharp MX-3501N'):
    printerComboBox.current(0)
printerComboBox.grid(column=0, row=1, sticky=tk.W, padx=12, pady=5)

#############################
# Program information tab   #
#############################
programInfoFrame = tk.Frame(master=infoTab, bg=machineFrameColor, highlightbackground=buttonColor, highlightcolor=buttonColor, highlightthickness=4)

# License label
licInfoLabel = tk.Label(master=programInfoFrame, text="© Highlight Industries 2022", font=textFont, bg=machineFrameColor, fg=textColor, padx=5, pady=2)
licInfoLabel.pack()

# Program version label
versionLabel = tk.Label(master=programInfoFrame, text="Version :  v2.1", font=textFont, bg=machineFrameColor, fg=textColor, padx=5, pady=2)
versionLabel.pack()

# Info date label
infoDateLabel = tk.Label(master=programInfoFrame, text="Date :  1/20/2022", font=textFont, bg=machineFrameColor, fg=textColor, padx=5, pady=2)
infoDateLabel.pack()

#########################
# Buttons               #
#########################
# Done button
doneButton = tk.Button(master=macButtonFrame, text="Done", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.doneCallBack(macVarS, 
                                                                                                                                                                                                    custEntry, distrEntry, projNumEntry, manYearEntry,
                                                                                                                                                                                                    phaseEntry, mainLineVEntry, controlVEntry, totMotorEntry, fullLoadEntry,
                                                                                                                                                                                                    enginEntry, dateEntry,
                                                                                                                                                                                                    synVar, profileVar, scaleVar, coldVar, AMP20V, splitVar, autoVar, ulVar, doorVar, harVar,
                                                                                                                                                                                                    numRan))
doneButton.grid(column=0, row=0, sticky=tk.E, padx=5, pady=5)

# Main Stop button
stopButton = tk.Button(master=macButtonFrame, text="Stop", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.stopCallBack(projNumEntry))
stopButton.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)

# Reset button
resetButton = tk.Button(master=macButtonFrame, text="Reset", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command=lambda: callBack.resetCallBack(macVarS, synVar, profileVar, scaleVar, coldVar, AMP20V, splitVar, autoVar, ulVar, doorVar, harVar,
                                                                                                                                                                                                    custEntry, distrEntry, projNumEntry, manYearEntry, enginEntry, dateEntry,
                                                                                                                                                                                                    phaseEntry, mainLineVEntry, controlVEntry, totMotorEntry, fullLoadEntry))
resetButton.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

# Export button
exButton = tk.Button(master=exButtonFrame, text="Export to Excel", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.exportCallBack(exJobNumEntry, exDirEntry, exportJobVar, numRan))
exButton.grid(column=0, row=0, sticky=tk.E, padx=5, pady=5)

# Export stop button
exStopButton = tk.Button(master=exButtonFrame, text="Stop", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.exStopCallBack(exJobNumEntry))
exStopButton.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)

# Access button
accessButton = tk.Button(master=accessButtonFrame, text="Import to Access", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.accessCallBack(accessJobEntry, accessDirEntry, accessJobVar))
accessButton.grid(column=0, row=0, padx=5, pady=5)

# Print button
printButton = tk.Button(master=printButtonFrame, text="Print", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.printCallBack(dirEntry, jobEntry, pagesEntry, printerVar, jobVar, paVar, numRan))
printButton.grid(column=0, row=0, sticky=tk.E, padx=5, pady=5)

# Print Stop button
printStopButton = tk.Button(master=printButtonFrame, text="Stop", font=miniFont, activebackground=machineFrameColor, activeforeground=textColor, bg=buttonColor, fg=machineFrameColor, command= lambda: callBack.printStopCallBack(jobEntry))
printStopButton.grid(column=1, row=0, sticky=tk.W, padx=5, pady=5)

# Frame Packing
machineIDFrame.pack(anchor="n", fill=tk.BOTH, side=tk.TOP)
infoFrame.pack(fill=tk.BOTH, side=tk.TOP)
machTypeFrame.pack(fill=tk.BOTH, side=tk.TOP)
printFrame.pack(fill=tk.BOTH, side=tk.TOP)
jobFrame.pack(fill=tk.BOTH, side=tk.TOP)
pageFrame.pack(fill=tk.BOTH, side=tk.TOP)
machStyleFrame.pack(fill=tk.BOTH, side=tk.TOP)
synFrame.pack(fill=tk.BOTH, side=tk.TOP)
opAddFrame.pack(fill=tk.BOTH, side=tk.TOP)
utilitySpecsFrame.pack(fill=tk.BOTH, side=tk.TOP)
exFrame.pack(fill=tk.BOTH, side=tk.TOP)
exportFrame.pack(fill=tk.BOTH, side=tk.TOP)
programInfoFrame.pack(fill=tk.BOTH, side=tk.TOP)
printerFrame.pack(fill=tk.BOTH, side=tk.TOP)
macButtonFrame.pack(fill=tk.BOTH, side=tk.BOTTOM)
printButtonFrame.pack(fill=tk.BOTH, side=tk.BOTTOM)
exButtonFrame.pack(fill=tk.BOTH, side=tk.TOP)
accessFrame.pack(fill=tk.BOTH, side=tk.TOP)
accessButtonFrame.pack(fill=tk.BOTH, side=tk.BOTTOM)
accessJobFrame.pack(fill=tk.BOTH, side=tk.TOP)

# Main window loop
window.mainloop()
