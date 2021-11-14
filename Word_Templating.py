import pandas as pd
import os
import sys
from docxtpl import DocxTemplate


#Enter the complete filepath to the Word Template
fname2=r"F:\WEBINARS\SEMINAR_02102021\4. Word Templating\appointment-letter-format_1.docx"


#Enter the folder path to where you want to save the final word  files
os.chdir(r"F:\WEBINARS\SEMINAR_02102021\4. Word Templating")


#Enter the complete filepath to the excel file which has the data
df=pd.read_excel(r"F:\WEBINARS\SEMINAR_02102021\4. Word Templating\Base Data_Candidates.xlsx")

Date=df["Date"].values
Candidate_Name=df["Candidate_Name"].values
Address_Line1=df["Address_Line1"].values
Address_Line2=df["Address_Line2"].values
City=df["City"].values
State=df["State"].values
Salut=df["Salut"].values
Designation=df["Designation"].values
Company_name=df["Company_name"].values
Joining_Date=df["Joining_Date"].values
Posting=df["Posting"].values
Supervisor_Name=df["Supervisor_Name"].values


zipped=zip(Date,Candidate_Name,Address_Line1,Address_Line2,City,State,Salut,Designation,Company_name,Joining_Date,Posting,Supervisor_Name)

for a,b,c,d,e,f,g,h,i,j,k,l in zipped:


    doc=DocxTemplate(fname2)

    context={"Date":a,"Candidate_Name":b,"Address_Line1":c,"Address_Line2":d,"City":e,"State":f,"Salut":g,"Designation":h,"Company name":i,"Joining_Date":j,"Posting":k,"Supervisor_Name":l}

    doc.render(context)
    doc.save('{}.docx'.format(b))
    
print("All Files done")