###Basic Understanding required to Understabd this code
## 1.Basic Concept of Variable, Loop, string Formatting , pandas read excel (Covered in Video -Part 3- Basics of Python)
## 2. Basic Concept of os & glob Module (Covered in Video -Part 5 - OS & GLOB Module)


import pandas as pd

import os
import glob

filepath="G:\OFFICE AUTOMATION FILES\Combine Excel Files\April.xlsx"

dirname=os.path.dirname(filepath)

files=glob.glob(os.path.join(dirname,"*.xlsx"))

df2=pd.DataFrame()

for f in files:
    df1=pd.read_excel(f,sheet_name=0)

    df2=df2.append(df1)
    #
    # print("the file is {}".format(f))

combfile=os.path.join(dirname,"combined_file.xlsx")

df2.to_excel(combfile,sheet_name="Combined")

