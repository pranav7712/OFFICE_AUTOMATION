# import modules


from tkinter import *
import pandas as pd
from selenium import webdriver
from datetime import date

date = date.today()
date = str(date)
import time
from openpyxl import load_workbook

# create window to use in app mode
fotagui = Tk()
fotagui.geometry("300x300")
fotagui.title("GSTN STATUS")

label_1 = Label(fotagui, text='Please enter no.of GSTNs to check')
label_1.pack()

e_1 = Entry(fotagui, width=5, bg="black", fg="white")
e_1.pack()


# define function to link with window's button

def program():
    # read excel to get input
    df = pd.read_excel("G:\Gst cHECK.xlsx", nrows=int(e_1.get()), engine='openpyxl')

    # define cols from which input to be taken
    cols = df["Gst Reg. No."].values

    # path for output
    path = r"G:\Output.xlsx"

    book = load_workbook(path)

    # use writer to  get output in excel file
    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book

    df = pd.DataFrame({"GSTIN": [], "STATUS": []})

    # define driver
    driver = webdriver.Chrome(executable_path=r"C:\Users\Dell\OneDrive - Indian Oil Corporation Limited\Documents\chromedriver.exe")
    # open particular website
    driver.get('https://www.mastersindia.co/gst-number-search-and-gstin-verification/')

    for i in cols:
        # input in searchbox
        searchbox = driver.find_element_by_xpath('//*[@id="gstin-search-form"]/div/div/input')

        searchbox.clear()
        # click search button
        searchbox.send_keys(i)

        searchbutton = driver.find_element_by_xpath('//*[@id="gstin-search-buton"]')
        searchbutton.click()
        time.sleep(2)

        output = driver.find_element_by_xpath('/html/body/main/section[2]/div/div[2]/div/table/tbody/tr/td[10]').text
        # s=[i,output]
        df = df.append({"GSTIN": i, "STATUS": output}, ignore_index=True)

    driver.quit()

    df

    df.to_excel(writer, sheet_name=date, index=False)
    # to save excel file
    writer.save()
    writer.close()

    label_2 = Label(fotagui, text="Process completed")
    label_2.pack(padx=50, pady=50)
    return


Button_1 = Button(fotagui, text="CLICK HERE", command=program)
Button_1.pack(padx=10, pady=10)

fotagui.mainloop()

