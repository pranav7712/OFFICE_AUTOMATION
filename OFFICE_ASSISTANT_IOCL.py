from tkinter import *
from tkinter import filedialog, messagebox , ttk
import pandas as pd
import os
import glob
from openpyxl import load_workbook
from shutil import copyfile
import datetime



FotaGui = Tk()

LogGui=Tk()

FotaGui.geometry('500x600')
FotaGui.title('Office Assistant for IOCL  ')

LogGui.geometry('800x400')
LogGui.title('Log of all activities:')

filepath = ''

label_head0 = Label(FotaGui, text="   Office Assistant for IOCL Users:"
                                   "\n Program for automating various office functions"
                                   "\n" , font="Times 10 ")

label_head0.pack()

label_head0 = Label(FotaGui, text="\n "
                                   "\n Click on HELP_INFO for more Information", font="Times 10 ",anchor=W)

label_head0.pack()




my_scrollbar1=Scrollbar(LogGui,orient=VERTICAL)
my_scrollbar1.pack(side=RIGHT,fill=Y)

my_scrollbar2=Scrollbar(LogGui,orient=HORIZONTAL)
my_scrollbar2.pack(side=BOTTOM,fill=Y)





def HELP_INFO():
    root=Tk()

    label_head0 = Label(root, text="   Program for office functions", font="Times 10 ")

    label_head0.pack()

    label_head1 = Label(root, text='This is the program to perform various tasks in Pdf or Text or  Excel Files in 2 steps .'
                                      ' \n   .', bd=1, relief='solid', font='Times 12', anchor=N)
    label_head1.pack()

    label_head2 = Label(root, text='Step 1: Click "Browse" button to select the file which you want to perform action .'
                                   '\n The "Clear Memory" button can be used to clear the files selected and reselect the file',
                        bd=1, relief='solid', font='Times 12', anchor=N)
    label_head2.pack()

    label_head3 = Label(root, text='Step 2: Click on the action button at the below:', bd=1, relief='solid',
                        font='Times 12', anchor=N)
    label_head3.pack()

    label_head4 = Label(root,
                        text='\n "SPLIT FILES" In case you want to split files, you need to specify:'
                             '\n The exact name of the column based on which you want to split the file.'
                             '\n The Files can either be splitted into various sheets in same file or in different files'
                             ,
                        bd=1, relief='solid', font='Times 12', anchor=N
                        )

    label_head4.pack()

    label_head5 = Label(root, text='NOTE 2: For Combining Excel/pdf/Text files, please note below:'
                                      '\n All Excel or pdf or text files to be combined should be kept in a single folder'
                                      '\n On clicking browse, multiple files cannot be selected.'
                                      '\n So, only one file in that folder to be selected.'
                                      '\n The program will automatically read all other pdf files/ excel/ text files and merge them'
                                   '\n "COMBINE NORMAL EXCEL FILES":In case of Combining Excel, there are two options '
                                   '\n 1. Combine Files : In this The first sheet of all the Excel files will be merged.'
                                   '\n 2. Combine Sheets: In case , User has to select 1 excel file and then all the sheets of this excel file will be combined '
                                   '\n'
                                   '\n "Combine GSTR2A file": It is to be used specificaly for the GSTR 2A downloaded from GST Site (For Tax People Specifically)'
                                   '\n'
                                   '\n "Combine PDF Files": This will merge the PDF files in the sequence they are stored in the folder'
                                   '\n'
                                   '\n "Combine SAP TXT Files": This will Combine the various Text Files. Please ensure that rows heading are same for all text files    '
                        , bd=1, relief='solid', font='Times 12', anchor=N

                        )

    label_head5.pack()

    label_head6 = Label(root, text='For Any Other Queries/ Issues/ Feedback: Please Contact'
                                   '\n Pranav P Tulshyan'
                                   '\n Email= pratiktp@indianoil.in'
                                   '\n Mob: 9205526726'
                                   '\n'
                                   '\n We want feedbak on this program so that it can be further improved and enhabced to meet needs of more users', bd=1, relief='solid',
                        font='Times 11', anchor=N)
    label_head6.pack()

    root.title('Help/ Info About the Program')

    root.mainloop()

label_head7=Label(LogGui)

labels=[]

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


def sendtofile(colslist, filepath):
    df = pd.read_excel(filepath)
    cols = e_1.get()
    pth = os.path.dirname(filepath)
    colslist = list(set(df[cols].values))
    global now



    for i in colslist:
        df[df[cols] == i].to_excel("{}/{}.xlsx".format(pth, i), sheet_name=i, index=False)

    messagebox.showinfo('Output', 'You data has been split into {} and {} files has been created.Click OK. \n All Files stored in same folder'.format(
                            ', '.join(colslist), len(colslist)))

    label_head7 = Label(LogGui,
                        text='{n}The Files have been Splitted to different Files.'.format(n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    print('\nCompleted')
    print('Thanks for using this program.')
    return


def sendtosheet(colslist):
    cols = e_1.get()
    extension = os.path.splitext(filepath)[1]
    filename = os.path.splitext(filepath)[0]
    pth = os.path.dirname(filepath)
    newfile = os.path.join(pth, filename + '_Sheet_Split_Auto' + extension)
    df = pd.read_excel(filepath)
    colslist = list(set(df[cols].values))



    copyfile(filepath, newfile)
    for j in colslist:
        writer = pd.ExcelWriter(newfile, engine='openpyxl')
        for myname in colslist:
            mydf = df.loc[df[cols] == myname]
            mydf.to_excel(writer, sheet_name=myname, index=False)
        writer.save()

    messagebox.showinfo('Output',
                        'You data has been split into {} and {} sheets has been created under single file named {new}.\n Click on OK .'.format(
                            ', '.join(colslist), len(colslist),new=newfile))

    label_head7 = Label(LogGui,
                        text='{n}The Files have been Splitted to different sheets.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    print('\nCompleted')
    print('Thanks for using this program.')
    return


def SPLIT_FILE():
    global filepath
    global e_1
    global e_2




    splitwin=Tk()

    label_1 = Label(splitwin, text='Enter the Exact Column name whose value u want to Split')
    label_1.pack()
    e_1 = Entry(splitwin, width=50, bg='blue', fg='white', borderwidth=4)
    e_1.pack()



    Browsebutton = Button(splitwin, width=20, text="Split Files", command=SPLIT_FILE2)
    Browsebutton.pack()

    splitwin.mainloop()

def SPLIT_FILE2():

    df = pd.read_excel(filepath)
    cols = e_1.get()
    colslist = list(set(df[cols].values))

    messagebox.showinfo('Check the output',
                        'You data will split based on these values {} and create {} files or sheets based on next selection. If you are ready to proceed Click OK or close the dialog box to re-start.'.format(
                            ', '.join(colslist), len(colslist)))

    response=messagebox.askyesno('Split Files','Do you want to split in Various Sheets in Same file  OR Different Files? '
                                               '\nClick Yes for Various Sheets in Same File.!'
                                               '\n CLick No For Different Files')
    df = pd.read_excel(filepath)
    cols = e_1.get()
    colslist = list(set(df[cols].values))
    
    if response == 0:
        sendtofile(colslist, filepath)
    elif response == 1:
        sendtosheet(colslist)
    else:
        messagebox.showerror('Output',"Something went wrong")


def Combine_File():
    global filepath
    global sheet_pos
    combwin = Tk()
    combwin.title('Combining Excel files')



    label_comb1 = Label(combwin, text='Note: There are two options for combining the Excel file'
                                      '\n '
                                      '\n Option 1: Combine Files--> This will combine the first sheet of all the excel files'
                                      '\n Option 2: Combine Sheets--> This will Combine all the sheets of that particular excel file'
                                      '\n'
                                      '\n While choosing option 2, please ensure that u have only that 1 excel file (whose sheets are to be combined) in that folder.')
    label_comb1.pack()

    label_1 = Label(combwin, text='\n'
                                  '\n')
    label_1.pack()

    Browsebutton = Button(combwin, width=15, text="Opt1:Combine Files", command=Combine_File2)
    Browsebutton.pack()

    Browsebutton = Button(combwin, width=15, text="Opt2:Combine Sheets", command=Combine_File3)
    Browsebutton.pack()

    label_1 = Label(combwin, text='\n'
                                  '\n')
    label_1.pack()

    combwin.mainloop()


def Combine_File2():
    global filepath
    global sheet_pos


    pth = os.path.dirname(filepath)
    extension = os.path.splitext(filepath)[1]
    files = glob.glob(os.path.join(pth, '*.xls*'))
    newfile = os.path.join(pth, 'All_Files_Combined_Auto.xlsx')
    df = pd.DataFrame()
    response=messagebox.askyesno('Important Check','It has to be ensured that the Row Heading for All the excel files to be combined '
                                          'are exactly same.\n Any Difference in Row Heading may not combine Excel Properly.!'
                                          '\n Do You want to Continue?')


    if response==1:

        for f in files:

            data = pd.read_excel(f,sheet_name=0)
            df = df.append(data)

        df.to_excel(newfile, sheet_name='combined', index=False)

        messagebox.showinfo('Output', 'All excel files in the selected folder have been combined.\n Click on OK')
        label_head7 = Label(LogGui,
                            text='{n}The Excel Files were Combined.'.format(
                                n=now.strftime('%y-%m-%d %H:%M:%S')),
                            bd=1, relief='solid',
                            font='Times 10', anchor=N)
        label_head7.pack()
    else:
        messagebox.showinfo('Output','Operation terminated. PLease come back after alinging Row Heading Names. \n Click on OK')
        label_head7 = Label(LogGui,
                            text='{n}The operation was terminated.'.format(
                                n=now.strftime('%y-%m-%d %H:%M:%S')),
                            bd=1, relief='solid',
                            font='Times 10', anchor=N)
        label_head7.pack()


def Combine_File3():
    global filepath
    global sheet_pos


    pth = os.path.dirname(filepath)

    df = pd.DataFrame()

    df2 = pd.DataFrame()

    xl = pd.ExcelFile(filepath)


    newfile = os.path.join(pth, 'All_Sheets_Combined_Auto.xlsx')

    response=messagebox.askyesno('Important Check','It has to be ensured that the Row Heading for All the excel sheets to be combined '
                                          'are exactly same.\n Any Difference in Row Heading may not combine Excel Properly.!'
                                          '\n Do You want to Continue?')


    if response==1:
        res = len(xl.sheet_names)

        while res>0:
            res-=1
            df=pd.read_excel(filepath,sheet_name=res)
            df2=df2.append(df)

        df2.to_excel(newfile, sheet_name='combined', index=False)

        messagebox.showinfo('Output', 'All excel Sheets in the selected Excel File have been combined. \n Click on OK')
        label_head7 = Label(LogGui,
                            text='{n}The Excel Sheets were Combined.'.format(
                                n=now.strftime('%y-%m-%d %H:%M:%S')),
                            bd=1, relief='solid',
                            font='Times 10', anchor=N)
        label_head7.pack()
    else:
        messagebox.showinfo('Output','Operation terminated. PLease come back after alinging Row Heading Names.\n Click on OK')
        label_head7 = Label(LogGui,
                            text='{n}The operation was terminated.'.format(
                                n=now.strftime('%y-%m-%d %H:%M:%S')),
                            bd=1, relief='solid',
                            font='Times 10', anchor=N)
        label_head7.pack()




def Combine_PDF():
    import glob
    import os
    global filepath
    global label_head7
    global now


    from PyPDF2 import PdfFileMerger



    pth = os.path.dirname(filepath)

    filenames = glob.glob(pth + "/*.pdf")

    merged = PdfFileMerger()

    for files in filenames:
        merged.append(files)

    extension = os.path.splitext(filepath)[1]
    filename = os.path.splitext(filepath)[0]
    pth = os.path.dirname(filepath)
    newfile = os.path.join(pth, 'Combined_Pdf_File_Auto' + extension)

    merged.write(newfile)
    merged.close()

    messagebox.showinfo('Output', 'All pdf files in the selected folder have been merged.\n Click on OK')
    label_head7 = Label(LogGui,
                        text='{n}The PDF Files have been Combined.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()


def Combine_SAP_Txt_files():
    import pandas as pd
    import glob
    global filepath
    import os

    pth = os.path.dirname(filepath)
    extension = os.path.splitext(filepath)[1]

    newfile = os.path.join(pth, 'Combined_Text_File_Auto.txt')

    filenames = glob.glob(pth + "/*.txt")

    df2 = pd.DataFrame()

    for files in filenames:
        df = pd.read_csv(files, sep="\t", low_memory=False,encoding='cp1252')

        df2 = df2.append(df)

    df2.to_csv(newfile, sep="\t")

    messagebox.showinfo('Output', 'All text files in the selected folder have been merged.\n Click on OK')
    label_head7 = Label(LogGui,
                        text='{n}The SAP Text Files have been Combined.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    print('All SAP text files in the selected folder have been merged')


def Combine_GSTR2A_File():
    import pandas as pd
    import glob
    import os
    global filepath

    pth = os.path.dirname(filepath)



    filenames = glob.glob(pth + "/*.xlsx")



    i = 0
    for file in filenames:
        i = i + 1

    if i < 1:
        print("Upload at least 2 files")
    elif i >60:
        print("Maximum capacity is 60 files at a time")
    else:
        pass


    cum_size = 0

    for file in filenames:
        size = os.path.getsize(file)

        cum_size = cum_size + size

        if size > 31457280:
            print("Please upload a smaller file size. Maximum limit is 30 mb.")

        elif cum_size > 314572800:
            print("Combined File size for all the file is more than 300 mb. Please use smaller files")
            break
        else:
            pass

    # A. iterate through each file to append it one below the other

    # A.1 : This will iterate through the B2B file

    df2 = pd.DataFrame()

    for files in filenames:
        df = pd.read_excel(files, sheet_name=1)

        df1 = df.drop([0, 1, 2, 3, 4], axis=0)

        df1 = df1.dropna(how='all')

        df1['File_name'] = files

        df2 = df2.append(df1)

    # this is used for deleting all the rows which are totally blank

    df3 = df2

    # this is used for renaming the names of the columns

    df3.rename(columns={'Goods and Services Tax  - GSTR 2A': 'GSTIN_of_Supplier'}, inplace=True)
    df3.rename(columns={'Unnamed: 1': 'Legal_Name_Of Supplier'}, inplace=True)
    df3.rename(columns={'Unnamed: 2': 'Inv_CN_DN_Number_Original'}, inplace=True)
    df3.rename(columns={'Unnamed: 3': 'Inv_CN_DN_Type_Original'}, inplace=True)
    df3.rename(columns={'Unnamed: 4': 'Inv_CN_DN_Date_Original'}, inplace=True)
    df3.rename(columns={'Unnamed: 5': 'Inv_CN_DN_Value_Original'}, inplace=True)
    df3.rename(columns={'Unnamed: 6': 'Place_Of_Supply'}, inplace=True)
    df3.rename(columns={'Unnamed: 7': 'Supply_Attract_Reverse_Charge'}, inplace=True)
    df3.rename(columns={'Unnamed: 8': 'GST_Rate'}, inplace=True)
    df3.rename(columns={'Unnamed: 9': 'Taxable_Value_Rs'}, inplace=True)
    df3.rename(columns={'Unnamed: 10': 'IGST_Rs'}, inplace=True)
    df3.rename(columns={'Unnamed: 11': 'CGST_Rs'}, inplace=True)
    df3.rename(columns={'Unnamed: 12': 'SGST_Rs'}, inplace=True)
    df3.rename(columns={'Unnamed: 13': 'Cess'}, inplace=True)
    df3.rename(columns={'Unnamed: 14': 'GSTR_1_5_Filing_Status'}, inplace=True)
    df3.rename(columns={'Unnamed: 15': 'GSTR_1_5_Filing_Date'}, inplace=True)
    df3.rename(columns={'Unnamed: 16': 'GSTR_1_5_Filing_Period'}, inplace=True)
    df3.rename(columns={'Unnamed: 17': 'GSTR_3B_Filing_Status'}, inplace=True)
    df3.rename(columns={'Unnamed: 18': 'Amendment_made_if_any'}, inplace=True)
    df3.rename(columns={'Unnamed: 19': 'Tax_Period_in_which_Amended'}, inplace=True)
    df3.rename(columns={'Unnamed: 20': 'Effective_date_of_cancellation'}, inplace=True)
    df3.rename(columns={'Unnamed: 21': 'Source'}, inplace=True)
    df3.rename(columns={'Unnamed: 22': 'IRN'}, inplace=True)
    df3.rename(columns={'Unnamed: 23': 'IRN_Date'}, inplace=True)

    # here we will remove the rows, in which the invoice number has  a total
    filt = df3['Inv_CN_DN_Number_Original'].str.contains('Total', na=False)
    df3 = df3[~filt]

    df3['Inv_CN_DN_Date_Unique'] = df3['Inv_CN_DN_Date_Original'].str.replace("-", ".")
    df3['Total_tax'] = df3['IGST_Rs'] + df3['CGST_Rs'] + df3['SGST_Rs']
    df3['Unique_ID'] = df3['GSTIN_of_Supplier'] + "/" + df3['Inv_CN_DN_Number_Original'] + "/" + df3[
        'Inv_CN_DN_Date_Unique']

    df3['Sheet_Name'] = ("B2B")

    label_head21=Label(FotaGui,text='The B2B table is being combined... Please wait')

    # A.2 : This will iterate through the B2BA file


    label_head7 = Label(LogGui,
                        text='{n}Opertaion in progress..... Please wait. Sheet B2B Being read.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    df2 = pd.DataFrame()

    for files in filenames:
        df = pd.read_excel(files, sheet_name=2)

        df1 = df.drop([0, 1, 2, 3, 4, 5], axis=0)

        df1 = df1.dropna(how='all')

        df1['File_name'] = files

        df2 = df2.append(df1)

    # this is used for deleting all the rows which are totally blank
    df4 = df2.dropna(how='all')

    # this is used for renaming the names of the columns

    df4.rename(
        columns={'                                      Goods and Services Tax - GSTR-2A': 'Inv_CN_DN_Number_Original'},
        inplace=True)
    df4.rename(columns={'Unnamed: 1': 'Inv_CN_DN_Date_Original'}, inplace=True)
    df4.rename(columns={'Unnamed: 2': 'GSTIN_of_Supplier'}, inplace=True)
    df4.rename(columns={'Unnamed: 3': 'Legal_Name_Of Supplier'}, inplace=True)
    df4.rename(columns={'Unnamed: 4': 'Inv_CN_DN_Type_Revised'}, inplace=True)
    df4.rename(columns={'Unnamed: 5': 'Inv_CN_DN_Number_Revised'}, inplace=True)
    df4.rename(columns={'Unnamed: 6': 'Inv_CN_DN_Date_Revised'}, inplace=True)
    df4.rename(columns={'Unnamed: 7': 'Inv_CN_DN_Value_Revised'}, inplace=True)
    df4.rename(columns={'Unnamed: 8': 'Place_Of_Supply'}, inplace=True)
    df4.rename(columns={'Unnamed: 9': 'Supply_Attract_Reverse_Charge'}, inplace=True)
    df4.rename(columns={'Unnamed: 10': 'GST_Rate'}, inplace=True)
    df4.rename(columns={'Unnamed: 11': 'Taxable_Value_Rs'}, inplace=True)
    df4.rename(columns={'Unnamed: 12': 'IGST_Rs'}, inplace=True)
    df4.rename(columns={'Unnamed: 13': 'CGST_Rs'}, inplace=True)
    df4.rename(columns={'Unnamed: 14': 'SGST_Rs'}, inplace=True)
    df4.rename(columns={'Unnamed: 15': 'Cess'}, inplace=True)
    df4.rename(columns={'Unnamed: 16': 'GSTR_1_5_Filing_Status'}, inplace=True)
    df4.rename(columns={'Unnamed: 17': 'GSTR_1_5_Filing_Date'}, inplace=True)
    df4.rename(columns={'Unnamed: 18': 'GSTR_1_5_Filing_Period'}, inplace=True)
    df4.rename(columns={'Unnamed: 19': 'GSTR_3B_Filing_Status'}, inplace=True)
    df4.rename(columns={'Unnamed: 20': 'Effective_date_of_cancellation'}, inplace=True)
    df4.rename(columns={'Unnamed: 21': 'Amendment_made_if_any'}, inplace=True)
    df4.rename(columns={'Unnamed: 22': 'Original_tax_period_in_which_reported'}, inplace=True)

    # here we will remove the rows, in which the invoice number has  a total
    filt = df4['Inv_CN_DN_Number_Revised'].str.contains('Total', na=False)

    df4 = df4[~filt]

    df4['Inv_CN_DN_Date_Unique'] = df4['Inv_CN_DN_Date_Original'].str.replace("-", ".")
    df4['Total_tax'] = df4['IGST_Rs'] + df4['CGST_Rs'] + df4['SGST_Rs']
    df4['Unique_ID'] = df4['GSTIN_of_Supplier'] + "/" + df4['Inv_CN_DN_Number_Original'] + "/" + df4[
        'Inv_CN_DN_Date_Unique']

    df4['Sheet_Name'] = ("B2BA")

    label_head7 = Label(LogGui,
                        text='{n}Opertaion in progress..... Please wait. Sheet B2BA Being read.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    # A.3 : This will iterate through the CDNR file

    df2 = pd.DataFrame()

    for files in filenames:
        df = pd.read_excel(files, sheet_name=3)

        df1 = df.drop([0, 1, 2, 3, 4], axis=0)

        df1 = df1.dropna(how='all')

        df1['File_name'] = files

        df2 = df2.append(df1)

    # this is used for deleting all the rows which are totally blank
    df5 = df2.dropna(how='all')

    # this is used for renaming the names of the columns

    df5.rename(
        columns={'                                             Goods and Services Tax - GSTR-2A': 'GSTIN_of_Supplier'},
        inplace=True)
    df5.rename(columns={'Unnamed: 1': 'Legal_Name_Of Supplier'}, inplace=True)
    df5.rename(columns={'Unnamed: 2': 'Credit_Debit_Note_Original'}, inplace=True)
    df5.rename(columns={'Unnamed: 3': 'Inv_CN_DN_Number_Original'}, inplace=True)
    df5.rename(columns={'Unnamed: 4': 'Inv_CN_DN_Type_Original'}, inplace=True)
    df5.rename(columns={'Unnamed: 5': 'Inv_CN_DN_Date_Original'}, inplace=True)
    df5.rename(columns={'Unnamed: 6': 'Inv_CN_DN_Value_Original'}, inplace=True)
    df5.rename(columns={'Unnamed: 7': 'Place_Of_Supply'}, inplace=True)
    df5.rename(columns={'Unnamed: 8': 'Supply_Attract_Reverse_Charge'}, inplace=True)
    df5.rename(columns={'Unnamed: 9': 'GST_Rate'}, inplace=True)
    df5.rename(columns={'Unnamed: 10': 'Taxable_Value_Rs'}, inplace=True)
    df5.rename(columns={'Unnamed: 11': 'IGST_Rs'}, inplace=True)
    df5.rename(columns={'Unnamed: 12': 'CGST_Rs'}, inplace=True)
    df5.rename(columns={'Unnamed: 13': 'SGST_Rs'}, inplace=True)
    df5.rename(columns={'Unnamed: 14': 'Cess'}, inplace=True)
    df5.rename(columns={'Unnamed: 15': 'GSTR_1_5_Filing_Status'}, inplace=True)
    df5.rename(columns={'Unnamed: 16': 'GSTR_1_5_Filing_Date'}, inplace=True)
    df5.rename(columns={'Unnamed: 17': 'GSTR_1_5_Filing_Period'}, inplace=True)
    df5.rename(columns={'Unnamed: 18': 'GSTR_3B_Filing_Status'}, inplace=True)
    df5.rename(columns={'Unnamed: 19': 'Amendment_made_if_any'}, inplace=True)
    df5.rename(columns={'Unnamed: 20': 'Tax_Period_in_which_Amended'}, inplace=True)
    df5.rename(columns={'Unnamed: 21': 'Effective_date_of_cancellation'}, inplace=True)
    df5.rename(columns={'Unnamed: 22': 'Source'}, inplace=True)
    df5.rename(columns={'Unnamed: 23': 'IRN'}, inplace=True)
    df5.rename(columns={'Unnamed: 24': 'IRN_Date'}, inplace=True)

    # here we will remove the rows, in which the invoice number has  a total
    filt = df5['Inv_CN_DN_Number_Original'].str.contains('Total', na=False)

    df5 = df5[~filt]

    df5['Inv_CN_DN_Date_Unique'] = df5['Inv_CN_DN_Date_Original'].str.replace("-", ".")
    df5['Total_tax'] = df5['IGST_Rs'] + df5['CGST_Rs'] + df5['SGST_Rs']
    df5['Unique_ID'] = df5['GSTIN_of_Supplier'] + "/" + df5['Inv_CN_DN_Number_Original'] + "/" + df5[
        'Inv_CN_DN_Date_Unique']

    df5['Sheet_Name'] = ("CDNR")


    label_head7 = Label(LogGui,
                        text='{n}Opertaion in progress..... Please wait. Sheet CDNR Being read.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    # A.2 : This will iterate through the CDNRA file

    df2 = pd.DataFrame()

    for files in filenames:
        df = pd.read_excel(files, sheet_name=4)

        df1 = df.drop([0, 1, 2, 3, 4, 5], axis=0)

        df1 = df1.dropna(how='all')

        df1['File_name'] = files

        df2 = df2.append(df1)

    # this is used for deleting all the rows which are totally blank
    df6 = df2.dropna(how='all')

    # this is used for renaming the names of the columns

    df6.rename(columns={'                             Goods and Services Tax - GSTR2A': 'Credit_Debit_Note_Original'},
               inplace=True)
    df6.rename(columns={'Unnamed: 1': 'Inv_CN_DN_Number_Original'}, inplace=True)
    df6.rename(columns={'Unnamed: 2': 'Inv_CN_DN_Date_Original'}, inplace=True)
    df6.rename(columns={'Unnamed: 3': 'GSTIN_of_Supplier'}, inplace=True)
    df6.rename(columns={'Unnamed: 4': 'Legal_Name_Of Supplier'}, inplace=True)
    df6.rename(columns={'Unnamed: 5': 'Credit_Debit_Note_Revised'}, inplace=True)
    df6.rename(columns={'Unnamed: 6': 'Inv_CN_DN_Number_Revised'}, inplace=True)
    df6.rename(columns={'Unnamed: 7': 'Inv_CN_DN_Type_Revised'}, inplace=True)
    df6.rename(columns={'Unnamed: 8': 'Inv_CN_DN_Date_Revised'}, inplace=True)
    df6.rename(columns={'Unnamed: 9': 'Inv_CN_DN_Value_Revised'}, inplace=True)
    df6.rename(columns={'Unnamed: 10': 'Place_Of_Supply'}, inplace=True)
    df6.rename(columns={'Unnamed: 11': 'Supply_Attract_Reverse_Charge'}, inplace=True)
    df6.rename(columns={'Unnamed: 12': 'GST_Rate'}, inplace=True)
    df6.rename(columns={'Unnamed: 13': 'Taxable_Value_Rs'}, inplace=True)
    df6.rename(columns={'Unnamed: 14': 'IGST_Rs'}, inplace=True)
    df6.rename(columns={'Unnamed: 15': 'CGST_Rs'}, inplace=True)
    df6.rename(columns={'Unnamed: 16': 'SGST_Rs'}, inplace=True)
    df6.rename(columns={'Unnamed: 17': 'Cess'}, inplace=True)
    df6.rename(columns={'Unnamed: 18': 'GSTR_1_5_Filing_Status'}, inplace=True)
    df6.rename(columns={'Unnamed: 19': 'GSTR_1_5_Filing_Date'}, inplace=True)
    df6.rename(columns={'Unnamed: 20': 'GSTR_1_5_Filing_Period'}, inplace=True)
    df6.rename(columns={'Unnamed: 21': 'GSTR_3B_Filing_Status'}, inplace=True)
    df6.rename(columns={'Unnamed: 22': 'Amendment_made_if_any'}, inplace=True)
    df6.rename(columns={'Unnamed: 23': 'Original_tax_period_in_which_reported'}, inplace=True)
    df6.rename(columns={'Unnamed: 24': 'Effective_date_of_cancellation'}, inplace=True)

    # here we will remove the rows, in which the invoice number has  a total
    filt = df6['Inv_CN_DN_Number_Revised'].str.contains('Total', na=False)

    df6 = df6[~filt]

    df6['Inv_CN_DN_Date_Unique'] = df6['Inv_CN_DN_Date_Original'].str.replace("-", ".")
    df6['Total_tax'] = df6['IGST_Rs'] + df6['CGST_Rs'] + df6['SGST_Rs']
    df6['Unique_ID'] = df6['GSTIN_of_Supplier'] + "/" + df6['Inv_CN_DN_Number_Original'] + "/" + df6[
        'Inv_CN_DN_Date_Unique']

    df6['Sheet_Name'] = ("CDNRA")


    label_head7 = Label(LogGui,
                        text='{n}Opertaion in progress..... Please wait. Sheet CDNRA Being read.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()

    # Making a combined sheet with all merged

    df8 = df3.append(df4)

    df9 = df8.append(df5)

    df10 = df9.append(df6)

    # maiking a sheet with person who did not file the GSTR 1

    df11 = df10[df10['GSTR_1_5_Filing_Status'] == "N"]

    df12 = df10[(df10['Supply_Attract_Reverse_Charge'] == "Y") & (df10['GSTR_1_5_Filing_Status'] == "Y")]

    df13 = df10[(df10['Supply_Attract_Reverse_Charge'] == "N") & (df10['GSTR_1_5_Filing_Status'] == "Y") & (
                df10['Total_tax'] < 1)]

    df14 = df10[(df10['Supply_Attract_Reverse_Charge'] == "N") & (df10['GSTR_1_5_Filing_Status'] == "Y") & (
                df10['Total_tax'] >= 1)]

    # saving the file with the name "Combined"

    extension = os.path.splitext(filepath)[1]
    filename = os.path.splitext(filepath)[0]
    pth = os.path.dirname(filepath)
    newfile = os.path.join(pth, filename + 'GSTR2A_all_combined' + extension)


    writer = pd.ExcelWriter(newfile, engine='openpyxl')

    df3.to_excel(writer, sheet_name="B2B")

    df4.to_excel(writer, sheet_name="B2BA")

    df5.to_excel(writer, sheet_name="CDNR")

    df6.to_excel(writer, sheet_name="CDNRA")

    titles = list(df10.columns)

    titles[0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], \
    titles[10], titles[11], titles[12], titles[13], titles[14], titles[15], titles[16], titles[17], titles[18], titles[
        19], titles[20], titles[21], titles[22], titles[23], titles[24], titles[25], titles[26], titles[27], titles[28], \
    titles[29], titles[30], titles[31], titles[32], titles[33], titles[34], titles[35] = titles[24], titles[28], titles[
        0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], titles[
                                                                                             10], titles[11], titles[
                                                                                             12], titles[13], titles[
                                                                                             26], titles[25], titles[
                                                                                             21], titles[27], titles[
                                                                                             14], titles[15], titles[
                                                                                             16], titles[17], titles[
                                                                                             18], titles[19], titles[
                                                                                             20], titles[22], titles[
                                                                                             23], titles[29], titles[
                                                                                             30], titles[31], titles[
                                                                                             32], titles[33], titles[
                                                                                             34], titles[35]

    df10[titles].to_excel(writer, sheet_name="All_Combined")

    titles = list(df11.columns)

    titles[0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], \
    titles[10], titles[11], titles[12], titles[13], titles[14], titles[15], titles[16], titles[17], titles[18], titles[
        19], titles[20], titles[21], titles[22], titles[23], titles[24], titles[25], titles[26], titles[27], titles[28], \
    titles[29], titles[30], titles[31], titles[32], titles[33], titles[34], titles[35] = titles[24], titles[28], titles[
        0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], titles[
                                                                                             10], titles[11], titles[
                                                                                             12], titles[13], titles[
                                                                                             26], titles[25], titles[
                                                                                             21], titles[27], titles[
                                                                                             14], titles[15], titles[
                                                                                             16], titles[17], titles[
                                                                                             18], titles[19], titles[
                                                                                             20], titles[22], titles[
                                                                                             23], titles[29], titles[
                                                                                             30], titles[31], titles[
                                                                                             32], titles[33], titles[
                                                                                             34], titles[35]

    df11[titles].to_excel(writer, sheet_name="GSTR_1_Not Filed")

    titles = list(df12.columns)

    titles[0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], \
    titles[10], titles[11], titles[12], titles[13], titles[14], titles[15], titles[16], titles[17], titles[18], titles[
        19], titles[20], titles[21], titles[22], titles[23], titles[24], titles[25], titles[26], titles[27], titles[28], \
    titles[29], titles[30], titles[31], titles[32], titles[33], titles[34], titles[35] = titles[24], titles[28], titles[
        0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], titles[
                                                                                             10], titles[11], titles[
                                                                                             12], titles[13], titles[
                                                                                             26], titles[25], titles[
                                                                                             21], titles[27], titles[
                                                                                             14], titles[15], titles[
                                                                                             16], titles[17], titles[
                                                                                             18], titles[19], titles[
                                                                                             20], titles[22], titles[
                                                                                             23], titles[29], titles[
                                                                                             30], titles[31], titles[
                                                                                             32], titles[33], titles[
                                                                                             34], titles[35]

    df12[titles].to_excel(writer, sheet_name="GSTR_Filed_RCM_Yes")

    titles = list(df13.columns)

    titles[0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], \
    titles[10], titles[11], titles[12], titles[13], titles[14], titles[15], titles[16], titles[17], titles[18], titles[
        19], titles[20], titles[21], titles[22], titles[23], titles[24], titles[25], titles[26], titles[27], titles[28], \
    titles[29], titles[30], titles[31], titles[32], titles[33], titles[34], titles[35] = titles[24], titles[28], titles[
        0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], titles[
                                                                                             10], titles[11], titles[
                                                                                             12], titles[13], titles[
                                                                                             26], titles[25], titles[
                                                                                             21], titles[27], titles[
                                                                                             14], titles[15], titles[
                                                                                             16], titles[17], titles[
                                                                                             18], titles[19], titles[
                                                                                             20], titles[22], titles[
                                                                                             23], titles[29], titles[
                                                                                             30], titles[31], titles[
                                                                                             32], titles[33], titles[
                                                                                             34], titles[35]

    df13[titles].to_excel(writer, sheet_name="Tax_Zero_Cases")

    titles = list(df14.columns)

    titles[0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], \
    titles[10], titles[11], titles[12], titles[13], titles[14], titles[15], titles[16], titles[17], titles[18], titles[
        19], titles[20], titles[21], titles[22], titles[23], titles[24], titles[25], titles[26], titles[27], titles[28], \
    titles[29], titles[30], titles[31], titles[32], titles[33], titles[34], titles[35] = titles[24], titles[28], titles[
        0], titles[1], titles[2], titles[3], titles[4], titles[5], titles[6], titles[7], titles[8], titles[9], titles[
                                                                                             10], titles[11], titles[
                                                                                             12], titles[13], titles[
                                                                                             26], titles[25], titles[
                                                                                             21], titles[27], titles[
                                                                                             14], titles[15], titles[
                                                                                             16], titles[17], titles[
                                                                                             18], titles[19], titles[
                                                                                             20], titles[22], titles[
                                                                                             23], titles[29], titles[
                                                                                             30], titles[31], titles[
                                                                                             32], titles[33], titles[
                                                                                             34], titles[35]

    df14[titles].to_excel(writer, sheet_name="Working_Cases")

    writer.save()

    messagebox.showinfo('Output','All GSTR2A files have been combined!. \n Click on OK')

    label_head7 = Label(LogGui,
                        text='{n}The GSTR 2A  Files have been Combined.'.format(
                            n=now.strftime('%y-%m-%d %H:%M:%S')),
                        bd=1, relief='solid',
                        font='Times 10', anchor=N)
    label_head7.pack()


Browsebutton = Button(FotaGui, width=15, text="HELP_INFO", command=HELP_INFO)
Browsebutton.pack()


label_0 = Label(FotaGui, text='\n'
                               '\n Step: 1 Select the File by clicking Browse Button !!!' ,font='Times 11', anchor=N)
label_0.pack()
Browsebutton = Button(FotaGui, width=15, text="BROWSE", command=file_path)
Browsebutton.pack()


print(filepath)

# label_1 = Label(FotaGui, text='Enter the Column name whose value u want to Split')
# label_1.pack()
# e_1 = Entry(FotaGui, width=50, bg='blue', fg='white', borderwidth=4)
# e_1.pack()

# label_2 = Label(FotaGui, text='Whether you want to split in Sheets(S)/ Files(F) ?')
# label_2.pack()
# e_2 = Entry(FotaGui, width=50, bg='blue', fg='white', borderwidth=4)
# e_2.pack()

def Clear_Memory():
    messagebox.showinfo('Memory Clear','The file selected have been cleared from memory')
    now=datetime.datetime.now()
    label_head12=Label(LogGui,text='{n}:The file selected have been cleared from memory. You may browse file again '.format(n=now.strftime("%y-%m-%d %H:%M:%S")))
    label_head12.pack()

Browsebutton = Button(FotaGui, width=20, text="Clear Memory", command=Clear_Memory)
Browsebutton.pack()




label_20 = Label(FotaGui, text='\n'
                               '\n Step: 2 Click on the Action which you want to Perform !!!'
                               '\n',font='Times 11', anchor=N)
label_20.pack()

Browsebutton = Button(FotaGui, width=20, text="Split Excel Files", command=SPLIT_FILE)
Browsebutton.pack()


Browsebutton = Button(FotaGui, width=20, text="Combine Normal Excel Files", command=Combine_File)
Browsebutton.pack()

Browsebutton = Button(FotaGui, width=20, text="Combine Pdf Files", command=Combine_PDF)
Browsebutton.pack()

Browsebutton = Button(FotaGui, width=20, text="Combine Txt Files", command=Combine_SAP_Txt_files)
Browsebutton.pack()


Browsebutton = Button(FotaGui, width=20, text="Combine GSTR2A Files", command=Combine_GSTR2A_File)
Browsebutton.pack()



label_head11=Label(LogGui, text='Log of all Activities:',anchor=W)
label_head11.pack()


label_head12 = Label(FotaGui, text="   \n"
                                    "\n"
                                    "\n"
                                    "\n Feedback for improving the Program is sought."
                                    "\n Based on Feedback, program can be improved to Cater needs of Specific Users"
                                    "\n Send your feedback at pratiktp@indianoil.in ", font="Times 10 ")

label_head12.pack()

LogGui.mainloop()

FotaGui.mainloop()