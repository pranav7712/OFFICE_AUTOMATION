import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
import os
import datetime

FotaGui = Tk()
LogGui=Tk()

FotaGui.geometry("500x500")
LogGui.geometry("400x400")


oldfile = StringVar()
newfile = StringVar()


def old_file():
    global oldfile
    global label_head7

    global now

    now = datetime.datetime.now()

    # Fetch the file path of the hex file browsed.
    if (oldfile == ""):
        oldfile = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select a file", filetypes=[("All Files", "*.*"),("Pdf Files","*.pdf"),("Text Files","*.txt"),("Excel FIles","*.xlsx")])
    else:
        oldfile = filedialog.askopenfilename(initialdir=oldfile,
                                              title="Select a file", filetypes=[("All Files", "*.*"),("Pdf Files","*.pdf"),("Text Files","*.txt"),("Excel FIles","*.xlsx")])

    extension = os.path.splitext(oldfile)[1]
    filename = os.path.splitext(oldfile)[0]
    pth = os.path.dirname(oldfile)
    # files = glob.glob(os.path.join(pth, '*{ext}'.format(ext=extension)))

    #
    #
    # for f in files:
    #
    label_head7 = Label(LogGui, text='{n}The File {fil} have been selected.'.format(fil=oldfile,n=now.strftime('%y-%m-%d %H:%M:%S')),bd=1, relief='solid',
                    font='Times 10', anchor=N)
    label_head7.pack()
    #


def new_file():

    global now
    global newfile


    now = datetime.datetime.now()



    # Fetch the file path of the hex file browsed.
    if (newfile == ""):
        newfile = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select a file", filetypes=[("All Files", "*.*"),("Pdf Files","*.pdf"),("Text Files","*.txt"),("Excel FIles","*.xlsx")])
    else:
        newfile = filedialog.askopenfilename(initialdir=newfile,
                                              title="Select a file", filetypes=[("All Files", "*.*"),("Pdf Files","*.pdf"),("Text Files","*.txt"),("Excel FIles","*.xlsx")])

    # extension = os.path.splitext(filepath)[1]
    # filename = os.path.splitext(filepath)[0]
    # pth = os.path.dirname(filepath)
    # files = glob.glob(os.path.join(pth, '*{ext}'.format(ext=extension)))
    #
    #
    #
    # for f in files:
    #
    label_head7 = Label(LogGui, text='{n}The File {fil} have been selected.'.format(fil=newfile,n=now.strftime('%y-%m-%d %H:%M:%S')),bd=1, relief='solid',
                    font='Times 10', anchor=N)
    label_head7.pack()

    #


def compare_files():
    global oldfile
    global newfile

    df1=pd.read_excel(oldfile)
    df2=pd.read_excel(newfile)


    print(df1)
    print(df2)

    print(df1.equals(df2))

    comparevalues=df1.values==df2.values

    print(comparevalues)


    (a,b)=np.where(comparevalues==False)
    print(a,b)

    zipped=list(zip(a,b))

    for (i,j) in zipped:
        df1.iloc[i,j]="{} is now {}".format(df1.iloc[i,j],df2.iloc[i,j])

    dir=os.path.dirname(oldfile)
    newfile=os.path.join(dir,"Compared.xlsx")
    df1.to_excel(newfile, index=False)


def compare_files():
    global oldfile
    global newfile

    df1=pd.read_excel(oldfile)
    df2=pd.read_excel(newfile)


    print(df1)
    print(df2)

    print(df1.equals(df2))

    comparevalues=df1.values==df2.values

    print(comparevalues)


    (a,b)=np.where(comparevalues==False)
    print(a,b)

    zipped=list(zip(a,b))

    for (i,j) in zipped:
        df1.iloc[i,j]="{} is now {}".format(df1.iloc[i,j],df2.iloc[i,j])


    dir=os.path.dirname(oldfile)
    comfile=os.path.join(dir,"Compared.xlsx")
    df1.to_excel(comfile, index=False)


def highlight_changes():
    global oldfile
    global newfile

    df1 = pd.read_excel(oldfile)
    df2 = pd.read_excel(newfile)

    comparevalues = df1.values == df2.values


    (a, b) = np.where(comparevalues == False)
    print(a, b)

    zipped = list(zip(a, b))

    for (i, j) in zipped:
        df1.iloc[i, j] = df2.iloc[i, j]
        df1.iloc[i, j] = df1.style.apply(lambda x: "background:red",axis=0)

    dir = os.path.dirname(oldfile)
    newfile = os.path.join(dir, "Highlighted.xlsx")
    df1.to_excel(newfile, index=False)


label_0 = Label(FotaGui, text='\n')
label_0.pack()

label_0 = Label(FotaGui, text='Step: 1 Select the File by clicking Browse Button !!!' ,font='Times 11', anchor=N,bd=1, relief='solid')
label_0.pack()

Browsebutton = Button(FotaGui, width=15, text="BROWSE_old_File", command=old_file)
Browsebutton.pack()


Browsebutton = Button(FotaGui, width=15, text="BROWSE_new_file", command=new_file)
Browsebutton.pack()


label_head3 = Label(FotaGui, text='\n'
                               '\n'
                               )
label_head3.pack()

label_head4 = Label(FotaGui, text='Step 2: Click on the action button at the below:', bd=1, relief='solid',
                        font='Times 12', anchor=N)

label_head4.pack()

Button_1=Button(FotaGui,text="Compare_Excel",command=compare_files)
Button_1.pack()


label_head12 = Label(FotaGui, text="   \n"
                                    "\n"
                                    "\n"
                                    "\n Feedback for improving the Program is sought."
                                    "\n Based on Feedback, program can be improved to Cater needs of Specific Users"
                                    "\n Send your feedback at pranav.tulshyan@gmail.com ", font="Times 10 ")

label_head12.pack()




LogGui.mainloop()
FotaGui.mainloop()
