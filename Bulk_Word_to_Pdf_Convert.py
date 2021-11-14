from docx2pdf import convert
import os
import glob


#Enter the Folder path which has all the word files for converting to pdf

path=r"F:\WEBINARS\SEMINAR_02102021\Word to Pdf Bulk"

files=glob.glob(path+"/*.docx*")

for f in files:
    convert(f)

print("All files converted")