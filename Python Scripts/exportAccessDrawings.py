import csv
import pypyodbc
import os
import tkinter as tk
from tkinter import *
from tkinter import messagebox

#
# converting from csv to access database
#
def toAccess(mainDir, jobNum):
    # Switching to access data directory
    mainDir = mainDir.replace("Drawings", "Electrical BOM")
    mainDir = mainDir + "\\"
    
    # Check to see if BOM folder exists
    if os.path.exists(mainDir) == False:
        tk.messagebox.showerror(title=jobNum + " Electrical BOM folder not found", message=jobNum + " Electrical BOM folder was not found and will be created please fill with " + jobNum + "_Table29a.csv, " + jobNum + "_Table29b.csv, " + jobNum + "_Table30.csv, and " + jobNum + "_Table32.csv files.")
        os.mkdir(mainDir)
        return

    # Check to see if csv files exists
    testFile29a = mainDir + jobNum + "_Table29a.csv"
    testFile29b = mainDir + jobNum + "_Table29b.csv"
    testFile30 = mainDir + jobNum + "_Table30.csv"
    testFile32 = mainDir + jobNum + "_Table32.csv"
    if os.path.isfile(testFile29a) == False or os.path.isfile(testFile29b) == False or os.path.isfile(testFile30) == False or os.path.isfile(testFile32) == False:
        tk.messagebox.showerror(title=jobNum + " Electrical BOM folder empty or wrong file name", message=jobNum + " Electrical BOM folder was empty or filled with incorrectly named files, please fill with " + jobNum + "_Table29a.csv, " + jobNum + "_Table29b.csv, " + jobNum + "_Table30.csv, and " + jobNum + "_Table32.csv files.")
        return
    
    # initializing highlight db file
    highDBFile = "F:\\SHARED\\SQL Elec Engineering\\accdb format files\\SQLEE8Austin.accdb"

    # connection string
    connstr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(highDBFile)

    # connecting to db
    conn = pypyodbc.connect(connstr)

    # creating cursor
    curs = conn.cursor()

    # sql query to clear job bill board table
    curs.execute("""DELETE FROM JobBillBoard;""")

    # set tableEMach to tbale 29b csv
    tabelEMach = mainDir + '\\' + jobNum + '_Table29b.csv'

    # qry to input data into access
    qry = "INSERT INTO JobBillBoard ([Job#], [MachineSection], [Part#], [QtyToUseInJob]) VALUES (?, ?, ?, ?)"

    # table 29b opening csv
    with open(tabelEMach) as eMachCSV:
        dataLines = csv.reader(eMachCSV, delimiter=',')
        lineCount = 0
        for row in dataLines:
            if lineCount == 0:
                lineCount += 1
            else:
                if (row[1] != '-' and row[2] != '-'):
                    partNum = row[2].replace("'", "")
                    qty = row[1].replace("'", "")
                    newRow = [jobNum, 'EMach', partNum, qty]
                    curs.execute(qry, newRow)
                    lineCount += 1

    # set tableE to table 29a csv
    tableE = mainDir + '\\' + jobNum + '_Table29a.csv'

    # table 29a opening csv
    with open(tableE) as eCSV:
        dataLines = csv.reader(eCSV, delimiter=',')
        lineCount = 0
        for row in dataLines:
            if lineCount == 0:
                lineCount += 1
            else:
                if (row[1] != '-' and row[2] != '-'):
                    partNum = row[2].replace("'", "")
                    qty = row[1].replace("'", "")
                    newRow = [jobNum, 'E', partNum, qty]
                    curs.execute(qry, newRow)
                    lineCount += 1

    # set tableE to table 30 csv
    tableE = mainDir + '\\' + jobNum + '_Table30.csv'

    # table 30 opening csv
    with open(tableE) as eCSV:
        dataLines = csv.reader(eCSV, delimiter=',')
        lineCount = 0
        for row in dataLines:
            if lineCount == 0:
                lineCount += 1
            else:
                if (row[1] != '-' and row[2] != '-'):
                    partNum = row[2].replace("'", "")
                    qty = row[1].replace("'", "")
                    newRow = [jobNum, 'E', partNum, qty]
                    curs.execute(qry, newRow)
                    lineCount += 1

    # set tableE to table 32 csv
    tableEMach = mainDir + '\\' + jobNum + '_Table32.csv'

    # table 32 opening csv
    with open(tableEMach) as eMachCSV:
        dataLines = csv.reader(eMachCSV, delimiter=',')
        lineCount = 0
        for row in dataLines:
            if lineCount == 0:
                lineCount += 1
            else:
                if (row[1] != '-' and row[2] != '-'):
                    partNum = row[2].replace("'", "")
                    qty = row[1].replace("'", "")
                    newRow = [jobNum, 'EMach', partNum, qty]
                    curs.execute(qry, newRow)
                    lineCount += 1

    # commit changes
    conn.commit()

    # close connection
    conn.close()

    # open accdb file that contains changes
    try:
        os.startfile(highDBFile)
    except:
        tk.messagebox.showerror(title="Error Opening file", message="Error when trying to open Highlight database file")
        return
    
    # message sent to user prompting them to refresh billboard in access to see changes
    tk.messagebox.showwarning(title="Refresh database", message="Changes have been made to the database, if access was previously open please hit 'Refresh BillBoard'")

