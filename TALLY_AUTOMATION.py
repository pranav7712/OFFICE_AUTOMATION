import pyautogui as pg
import time
import pandas as pd

FailSafeException=True


#OPENING THE TALLY APPLICATION


pg.click(x=2,y=1080,button="left")

time.sleep(1)

pg.write("Tally")


time.sleep(1)

pg.press("enter")

time.sleep(10)

pg.press("f1")

time.sleep(2)


pg.write("sabo")

time.sleep(2)


pg.press("down")

time.sleep(2)

pg.press("enter")

time.sleep(1)


pg.press("v")

time.sleep(2)

pg.press("f7")

time.sleep(2)


#for loopING THROUGH THE EXCEL FILES

akm=pd.read_excel(r"C:\Users\Desktop\Audit FY 21-22\Audit FY 21-22\Audit FY 21-22\Python\All Entries Row Added_1.xlsx")

jvno=akm["Transaction"].values


dates=akm["Date_1"].values

glname=akm["G_L Acct_BP Name"].values
amt=akm["amount"].values
drcr=akm["Debit/Credit"].values


zipped=zip(jvno,dates,glname,amt,drcr)

for (a,b,c,d,e) in zipped:
    
    if e=="Debit":
        time.sleep(0.1)
        
        astr=str(a)
        

        pg.write(astr)

        time.sleep(0.1)

        pg.press("enter")
        pg.write(b)

        time.sleep(0.1)
        pg.press("enter")

        pg.write(c)
        time.sleep(1)

        pg.press("enter")
        
        dstr=str(d)
        
        

        pg.write(dstr)
        time.sleep(1)


        pg.press("enter")
        
        pg.write("t")
        
        pg.press("enter")
        
    
    else:
        pg.write(c)
        time.sleep(1)
        
        pg.press("enter")
        
        dstr=str(d)
        pg.write(dstr)
        time.sleep(1)
        
        
        pg.press("enter")
        
        pg.press("enter")
        
        pg.press("enter")