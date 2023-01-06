Please add directory to AutoCADs trusted files to let the program run smoothly.
Follow steps:
1. Go to Options -> Files tab in AutoCAD
2. Expand trusten locations
3. Add "G:\~~EE Directory~~\01-AutoCAD Electrical\08-Standard Drawings Demo\Standard Machine Creator\AutoCAD Script\Synergy Semi Auto Scripts" to the paths

Additionally change AutoCad settings:
ATTDIA = 0
ATTREQ = 1
FILEDIA = 0
EXPERT = 2

Make sure all autoCAD tabs are closed or the program might catch an open drawing and run a script on it.

When exporting to excel the program will save tables as:
XXXXX_Table29a.csv
XXXXX_Table29b.csv
XXXXX_Table30.csv
XXXXX_Table32.csv

The program will turn off all object snap.