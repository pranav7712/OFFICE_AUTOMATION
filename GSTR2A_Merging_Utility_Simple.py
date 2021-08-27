#this version of GSTR2A merging utility will simply merge the columns of 4 sheets (B2B, B2BA, CDNR & CDNRA)
#No additional sheets / columns will be generated 

openpyxl import Workbook
from datetime import date
from tkinter import *
import os
import glob
from tkinter import messagebox,filedialog
import datetime
import warnings
import numpy as np
from UliPlot.XLSX import auto_adjust_xlsx_column_width


FotaGui=Tk()

LogGui=Tk()

FotaGui.geometry("500x500")
LogGui.geometry("250x250")

FotaGui.title("Utility for Merging GSTR2A")
LogGui.title("Log of all activities")

Label_1=Label(FotaGui,text="This is the utility for merging GSTR 2A",font="Times 16")
Label_1.pack()

warnings.filterwarnings('ignore')


def file_path():
    global filepath
    global label_head7
    filepath = StringVar()
    global now

    now = datetime.datetime.now()



    # Fetch the file path of the hex file browsed.
    if (filepath == ""):
        filepath = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select a file", filetypes=[("All Files", "*.*"),("Pdf Files","*.pdf"),("Text Files","*.txt"),("Excel FIles","*.xlsx")])
    else:
        filepath = filedialog.askopenfilename(initialdir=filepath,
                                              title="Select a file", filetypes=[("All Files", "*.*"),("Pdf Files","*.pdf"),("Text Files","*.txt"),("Excel FIles","*.xlsx")])

    extension = os.path.splitext(filepath)[1]
    filename = os.path.splitext(filepath)[0]
    pth = os.path.dirname(filepath)
    files = glob.glob(os.path.join(pth, '*{ext}'.format(ext=extension)))



    for f in files:

        label_head7 = Label(LogGui, text='{n}The File {fil} have been selected.'.format(fil=f,n=now.strftime('%y-%m-%d %H:%M:%S')),bd=1, relief='solid',
                        font='Times 10', anchor=N)
        label_head7.pack()




def Combine_GSTR2A_File():
    import pandas as pd
    import glob
    import os
    global filepath

    pth = os.path.dirname(filepath)



    filenames = glob.glob(pth + "/*.xlsx")



    df2=pd.DataFrame()
    
    for file in filenames:


        df=pd.read_excel(file,header=4,sheet_name ="B2B")

        df.rename(columns={'Invoice details': 'InvoiceNumber'}, inplace=True)
        df.rename(columns={'Unnamed: 3': 'Invoice Type'}, inplace=True)
        df.rename(columns={'Unnamed: 4': 'Invoice Date'}, inplace=True)
        df.rename(columns={'Unnamed: 5': 'Invoice Value'}, inplace=True)
        df.rename(columns={'Tax Amount': 'Integrated Tax  (₹)'}, inplace=True)
        df.rename(columns={'Unnamed: 11': 'Central Tax (₹)'}, inplace=True)
        df.rename(columns={'Unnamed: 12': 'State/UT tax (₹)'}, inplace=True)
        df.rename(columns={'Unnamed: 13': 'Cess  (₹)'}, inplace=True)




        df=df.drop(0)
        df=df.dropna(how='all')
        df = df[~df.InvoiceNumber.str.contains('Total')]

        df2=df2.append(df)


    df3=pd.DataFrame()

    for file in filenames:

        df4=pd.read_excel(file,header=5,sheet_name ="B2BA")

        df4.rename(columns={'Invoice details': 'Invoice Type'}, inplace=True)
        df4.rename(columns={'Unnamed: 5': 'InvoiceNumber2'}, inplace=True)
        df4.rename(columns={'Unnamed: 6': 'Invoice Date'}, inplace=True)
        df4.rename(columns={'Unnamed: 7': 'Invoice Value'}, inplace=True)
        df4.rename(columns={'Tax Amount': 'Integrated Tax  (₹)'}, inplace=True)
        df4.rename(columns={'Unnamed: 13': 'Central Tax (₹)'}, inplace=True)
        df4.rename(columns={'Unnamed: 14': 'State/UT tax (₹)'}, inplace=True)
        df4.rename(columns={'Unnamed: 15': 'Cess  (₹)'}, inplace=True)




        df4=df4.drop(0)
        df4=df4.dropna(how='all')
        df4 = df4[~df4.InvoiceNumber2.str.contains('Total')]

        df3=df3.append(df4)


    df6=pd.DataFrame()

    for file in filenames:   


        df5=pd.read_excel(file,header=4,sheet_name ="CDNR")
        df5.rename(columns={'Unnamed: 3': 'Notenumber'}, inplace=True)
        df5.rename(columns={'Unnamed: 4': 'Note Supply type'}, inplace=True)
        df5.rename(columns={'Unnamed: 5': 'Note date'}, inplace=True)
        df5.rename(columns={'Unnamed: 6': 'Note Value (₹)'}, inplace=True)
        df5.rename(columns={'Tax Amount': 'Integrated Tax  (₹)'}, inplace=True)
        df5.rename(columns={'Unnamed: 12': 'Central Tax (₹)'}, inplace=True)
        df5.rename(columns={'Unnamed: 13': 'State/UT tax (₹)'}, inplace=True)
        df5.rename(columns={'Unnamed: 14': 'Cess  (₹)'}, inplace=True)

        df5=df5.drop(0)
        df5=df5.dropna(how='all')
        df5 = df5[~df5.Notenumber.str.contains('Total')]

        df6=df6.append(df5)


    df7=pd.DataFrame()

    for file in filenames:  

        df8=pd.read_excel(file,header=5,sheet_name ="CDNRA")



        df8.rename(columns={'Unnamed: 6': 'NoteNumber'}, inplace=True)
        df8.rename(columns={'Unnamed: 7': 'Note Supply type'}, inplace=True)
        df8.rename(columns={'Unnamed: 8': 'Note date'}, inplace=True)
        df8.rename(columns={'Unnamed: 9': 'Note Value (₹)'}, inplace=True)
        df8.rename(columns={'Tax Amount': 'Integrated Tax  (₹)'}, inplace=True)
        df8.rename(columns={'Unnamed: 15': 'Central Tax (₹)'}, inplace=True)
        df8.rename(columns={'Unnamed: 16': 'State/UT tax (₹)'}, inplace=True)
        df8.rename(columns={'Unnamed: 17': 'Cess  (₹)'}, inplace=True)
        df8=df8.drop(0)
        df8=df8.dropna(how='all')
        df8 = df8[~df8.NoteNumber.str.contains('Total')]

        df7=df7.append(df8)


    q=date.today()
    filename=f'GSTR2A as on {q}'
    extension = os.path.splitext(filepath)[1]
    path=os.path.join(pth, filename + extension)




    writer=pd.ExcelWriter(path,engine='xlsxwriter', options={'strings_to_formulas': False})

    df2.to_excel(writer,sheet_name="B2B",index=False)
    auto_adjust_xlsx_column_width(df2, writer, sheet_name="B2B", margin=10)
    df3.to_excel(writer,sheet_name="B2BA",index=False)
    auto_adjust_xlsx_column_width(df2, writer, sheet_name="B2BA", margin=10)
    df6.to_excel(writer,sheet_name="CDNR",index=False)
    auto_adjust_xlsx_column_width(df2, writer, sheet_name="CDNR", margin=10)
    df7.to_excel(writer,sheet_name="CDNRA",index=False)
    auto_adjust_xlsx_column_width(df2, writer, sheet_name="CDNRA", margin=10)
    
    
    
    writer.save()
    writer.close()
    messagebox.showinfo('Output','All GSTR2A files have been combined!. \n Click on OK')
    
label_0 = Label(FotaGui, text='\n')
label_0.pack()

label_0 = Label(FotaGui, text='Step: 1 Select the File by clicking Browse Button !!!' ,font='Times 11', anchor=N,bd=1, relief='solid')
label_0.pack()

Browsebutton = Button(FotaGui, width=15, text="BROWSE", command=file_path)
Browsebutton.pack()


label_head3 = Label(FotaGui, text='\n'
                               '\n'
                               )
label_head3.pack()

label_head4 = Label(FotaGui, text='Step 2: Click on the action button at the below:', bd=1, relief='solid',
                        font='Times 12', anchor=N)

label_head4.pack()

Button_1=Button(FotaGui,text="Combine GSTR2A files",command=Combine_GSTR2A_File)
Button_1.pack()



LogGui.mainloop()

FotaGui.mainloop()