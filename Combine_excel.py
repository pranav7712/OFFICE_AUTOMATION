###Basic Concepts required to Understabd this code
## 1.Basic Concept of Variable, Loop, string Formatting , pandas read/write excel (Covered in Video -Part 3- Basics of Python)
## 2. Basic Concept of os & glob Module (Covered in Video -Part 5 - OS & GLOB Module)


##STEPS FOR Merging different sheets of Excel Files:

#1. Read excel file sheet 1
#2. Read excel file sheet 2 and so on
#3. On each reading, merge the data frame on a blank data Frame
#4 Save the Combined Excel file

import pandas as pd

filepath="G:\OFFICE AUTOMATION FILES\Combine Excel Files\Multiple Sheets\Multiple sheets.xlsx"

xl=pd.ExcelFile(filepath)

listt=xl.sheet_names

df2=pd.DataFrame()

for l in listt:
    df1=pd.read_excel(filepath,sheet_name=l)
    df2=df2.append(df1)

df2.to_excel("G:\OFFICE AUTOMATION FILES\Combine Excel Files\Multiple Sheets\Combined sheets.xlsx",sheet_name="Combined",index=False)




