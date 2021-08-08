#import the required modules

import pandas as pd
import os
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from datetime import date
import datetime
import time
import glob

#create the Tkinter main loop

fotagui = Tk()
fotagui.geometry("600x600")
fotagui.title("GSTIN STATUS REPORT")

#multiple labels to be displayed in the box

Label_0=Label(fotagui, text='This is the program to Check Multiple GSTIN Status at once .'
                                      ' \n   .', bd=1, relief='solid', font='Times 12', anchor=N)
Label_0.pack()

Label_1=Label(fotagui, text='You need to Upload a Input file using Browse Input File Button.'
                                      ' \n   .', bd=1, relief='solid', font='Times 12', anchor=N)
Label_1.pack()

Label_4=Label(fotagui,text="Format of Input Excel File"
                           "\n Simple mention the GSTIN in Column A"
                           "\n Keep Heading of Column A as GSTIN" , bd=1, relief='solid', font='Times 12', anchor=N )

Label_4.pack()
label_2 = Label(fotagui, text='\n'
                              '\n'
                              'Please write sets of GSTNs to check(Total no.of GSTNs/ 500) in Black Box Below:'
                              '\n For e.g if you have 1250 GSTIN in input file , then Mention 3'
                              '\n In case less than 500 GSTIN in input file, then Mention 1',bd=1, relief='solid', font='Times 12', anchor=N )
label_2.pack()

e_1 = Entry(fotagui, width=5, bg="black", fg="white")
e_1.pack()


#defining a function. This function will be attached to the button Browse File .

def browse_file():
    global fname
    fname = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("All files", "*")))


Label_5=Label(fotagui, text='Click Button below and Select that Excel file with GSTIN Numbers.'
                                      ' \n   .', bd=1, relief='solid', font='Times 12', anchor=N)
Label_5.pack()

x = Button(fotagui, text='Browse Input Excel File', command=browse_file)
x.pack()



#defining a function. This function will be attached to the button named Click Here.

def program():
    global fname
    df = pd.read_excel(fname, index_col=False)
    name = df.columns[0]
    df=df.rename(columns={name:"GSTIN"})  #this is done so that heading of col A is changed to GSTIN and is useful in merging

    dir_fname = os.path.dirname(fname)
    dir_fname1=dir_fname.replace("/", "\\")  #In selenium, while keeping default download path as Prefs, the one with / was showing error as it was not being read as a windows path. So changed that to \\
    print(dir_fname1)

    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": dir_fname1}   #this preference settings is done so that the file gets downloaded in the folder deaired andnot in the default download folder
    chrome_options.add_experimental_option('prefs', prefs)
    # driver = webdriver.Chrome(chrome_options=chrome_options)

    driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

    driver.get('https://my.gstzen.in/p/gstin-validator/')

    n = e_1.get()
    n = int(n)
    a = -500
    b = 0
    e = 1

    for i in range(n):
        c = df.iloc[a + 500:b + 500, 0:1]
        search = c.values

        searchbox = driver.find_element_by_xpath('//*[@id="id_text"]')

        searchbox.clear()

        searchbox.send_keys(str(search))

        searchbutton = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[4]/div/form/div[2]/div/button')
        searchbutton.click()
        time.sleep(2)

        output = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div[1]/h5/a')
        output.click()

        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')  #this creates a new tab , so no need to click back button
        driver.get('https://my.gstzen.in/p/gstin-validator/')

        a = a + 500
        b = b + 500

    # driver.quit()

    down_files=glob.glob(dir_fname1+"//TAXPAYER*.xlsx")   #this is toiterate through the downloaded files. Since the Downaloded files name start with TaxPaer, so we use glob to iterate all files with name starting with Taxpayer

    df2=pd.DataFrame()
    for i in down_files:
        df1=pd.read_excel(i)
        df2=df2.append(df1)  #this is for appending all the file and making a single data frame

    df4=df.merge(df2,left_on="GSTIN",right_on="GSTIN",how="left")  #here we are doing a Left Merge , so that the main file i.e df remains intact and other columns are addedin that file
    output_file = os.path.join(dir_fname1, "Output_GST.xlsx")
    df4.to_excel(output_file, engine="openpyxl" , index=False)

    messagebox.showinfo('Output', 'All GSTIN status have been extracted and kept in this location: {}!. \n Click on OK'.format(output_file))


Label_6=Label(fotagui, text='Click Button below to Generate Output file with GSTIN Status.'
                                      ' \n   .', bd=1, relief='solid', font='Times 12', anchor=N)
Label_6.pack()


Button_1 = Button(fotagui, text="CLICK to Check GSTINs", command=program)
Button_1.pack(padx=10, pady=10)

label_7 = Label(fotagui, text='\n'
                              '\n'
                              'For any query relating to the program, write to us at:'
                              '\n efficientcorporates.info@gmail.com'
                              '\n Feel free to provide your feedback & Suggestions.',bd=1, relief='solid', font='Times 12', anchor=N )
label_7.pack()



fotagui.mainloop()




