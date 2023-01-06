import csv
import msaccessdb
import pypyodbc
import shutil
import os

def toAccess(mainDir, jobNum):
    finalDB = (mainDir + '\\' + jobNum + '_Tables.accdb')
    msaccessdb.create(finalDB)

    # connection string
    connstr = "DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};".format(finalDB)

    # connecting to db
    conn = pypyodbc.connect(connstr)

    # creating cursor
    curs = conn.cursor()

    # creating new table
    curs.execute("""CREATE TABLE tableData (
                        ID              INT PRIMARY KEY NOT NULL,
                        ITEM            CHAR(100),
                        QTY             CHAR(100),
                        HLINum          CHAR(100),
                        MFG             CHAR(100),
                        CATALOG         CHAR(100),
                        DESCRIPTION     CHAR(100));""")
    
    # sql query to add data
    qry = "INSERT INTO tableData (ID, ITEM, QTY, HLINum, MFG, CATALOG, DESCRIPTION) VALUES (?, ?, ?, ?, ?, ?, ?)"
    
    # read data and send to access via sql query
    with open(jobNum + '_Tables.csv') as csvFile:
        dataLines = csv.reader(csvFile, delimiter=',')
        lineCount = 0
        for row in dataLines:
            if lineCount == 0:
                lineCount += 1
            else:
                # insert table data into db
                curs.execute(qry, row)
                lineCount += 1

    # delete blank lines from db
    curs.execute("""DELETE FROM tabledata 
                    WHERE QTY = '-'
                    OR HLINum = '-'
                    OR MFG = '-'
                    OR CATALOG = '-'
                    OR DESCRIPTION = '-';""")

    # commit changes
    conn.commit()

    # close database connection
    conn.close()

    # move files to job folder
    if not os.path.exists(mainDir + '\\' + jobNum + ' Drawings'):
        os.mkdir(mainDir + '\\' + jobNum + ' Drawings')

    shutil.move(mainDir + '\\' + jobNum + '_Tables.csv', mainDir + '\\' + jobNum + ' Drawings\\' + jobNum + '_Tables.csv')
    shutil.move(mainDir + '\\' + jobNum + '_Tables.xlsx', mainDir + '\\' + jobNum + ' Drawings\\' + jobNum + '_Tables.xlsx')
    shutil.move(mainDir + '\\' + jobNum + '_Tables.accdb', mainDir + '\\' + jobNum + ' Drawings\\' + jobNum + '_Tables.accdb')