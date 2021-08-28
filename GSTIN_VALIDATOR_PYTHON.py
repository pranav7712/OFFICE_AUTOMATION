from tkhtmlview import HTMLLabel
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
from PIL import ImageTk, Image
from ttkthemes import ThemedTk

fotagui = ThemedTk(theme="black")
fotagui.geometry("600x600")

fotagui.configure(background='black')

fotagui.title("GSTIN STATUS REPORT")

fotagui.minsize(600, 600)
fotagui.maxsize(600, 600)

my_canvas = Canvas(fotagui, width=600, height=600, bg='blue')
my_canvas.pack()

e_1 = Entry(fotagui, font=("Helventica", 18), width=8, fg="#336d92", bd=0, justify='center')
e_1.insert(0, "---")
entry_window = my_canvas.create_window(300, 300, window=e_1)


def entry_clear(e):
    if e_1.get() == '---':
        e_1.delete(0, END)


e_1.bind("<Enter>", entry_clear)


# e_1.pack()


def browse_file():
    global fname
    fname = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"), ("All files", "*")))


x = Button(fotagui, text='Browse Input Excel File', command=browse_file)
x_window = my_canvas.create_window(300, 200, window=x)


def program():
    global fname
    df = pd.read_excel(fname, index_col=False)
    name = df.columns[0]
    df = df.rename(columns={name: "GSTIN"})

    dir_fname = os.path.dirname(fname)
    dir_fname1 = dir_fname.replace("/", "\\")
    print(dir_fname1)

    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": dir_fname1}
    chrome_options.add_experimental_option('prefs', prefs)

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

        driver.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')
        driver.get('https://my.gstzen.in/p/gstin-validator/')

        a = a + 500
        b = b + 500

    driver.quit()

    down_files = glob.glob(dir_fname1 + "//TAXPAYER*.xlsx")

    df2 = pd.DataFrame()
    for i in down_files:
        df1 = pd.read_excel(i)
        df2 = df2.append(df1)

    df4 = df.merge(df2, left_on="GSTIN", right_on="GSTIN", how="left")
    output_file = os.path.join(dir_fname1, "Output_GST.xlsx")
    df4.to_excel(output_file, engine="openpyxl", index=False)

    messagebox.showinfo('Output',
                        'All GSTIN status have been extracted and kept in this location: {}!. \n Click on OK'.format(
                            output_file))


Button_1 = Button(fotagui, text="CLICK HERE", command=program)
Button_1_window = my_canvas.create_window(300, 370, window=Button_1)

my_label = HTMLLabel(fotagui,
                     html="<h1><a href='https://drive.google.com/drive/folders/1Ee-XzAVd-8N01_Ok0DhAzWJDZh447HKN?usp=sharing'>Learn to use!</a><h1>")

my_label_window = my_canvas.create_window(510, 600, window=my_label)

button6 = Button(fotagui, text="Help")
button7 = Button(fotagui, text="Exit", command=fotagui.destroy)

button6_window = my_canvas.create_window(20, 10, anchor="nw", window=button6)
button7_window = my_canvas.create_window(550, 10, anchor="nw", window=button7)

my_canvas.create_rectangle(0, 0, 600, 110, fill="#F0F0F0")

my_canvas.create_text(300, 60, text="This is the program to Check Multiple GSTIN Status at once .",
                      font=("Times New Roman", 15), fill="black")
my_canvas.create_text(300, 90, text="You need to Upload a Input file using Browse Input Excel File Button.",
                      font=("Times New Roman", 15), fill="black")
my_canvas.create_text(300, 150, text="               Format of Input Excel File:"
                                     "\n   Simple mention the GSTIN in Column A."
                                     "\n        Keep Heading of Column A as GSTIN.", font=("Bahnschrift Light", 12),
                      fill="white")

my_canvas.create_text(300, 250,
                      text="   Please write sets of GSTNs to check(Total no.of GSTNs/ 500) in White Box Below:"
                           "\n                 For e.g if you have 1250 GSTIN in input file , then Mention 3."
                           "\n                          In case less than 500 GSTIN in input file, then Mention 1.",
                      font=("Bahnschrift Light", 10), fill="white")

my_canvas.create_text(300, 335, text="Click Button below to Generate Output file with GSTIN Status.",
                      font=("Bahnschrift Light", 12), fill="white")

my_canvas.create_rectangle(0, 400, 600, 600, fill="#F0F0F0")

my_canvas.create_text(300, 387, text="For any issues / Feedbacks , Send mail at efficientcorporates.info@gmail.com",
                      font=("Bahnschrift Light", 10), fill="white")

fotagui.mainloop()





