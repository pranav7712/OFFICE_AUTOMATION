import glob
import os
import pyautogui
import time
from tkinter import messagebox

FailSafeException=True

files_to_sign = glob.glob(os.path.join("G:\OFFICE AUTOMATION FILES\Bills to Sign", '*.pdf*'))

for f in files_to_sign:
    time.sleep(2)
    pyautogui.click(x=1, y=1080, button="left")
    time.sleep(2)
    pyautogui.write("Acrobat Reader")
    pyautogui.press("enter")
    time.sleep(2)

    pyautogui.hotkey('ctrl','o')
    time.sleep(2)
    pyautogui.write(f)
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2    )
    pyautogui.click(x=127,y=77,button='left')  #click on tools
    time.sleep(2)
    pyautogui.click(x=363,y=535,button='left') #click on Certificates
    time.sleep(2)
    pyautogui.click(x=762,y=186,button='left') #click on Digitally Sign
    time.sleep(2)
    pyautogui.keyDown('ctrl')
    pyautogui.press('end')
    pyautogui.keyUp('ctrl')
    pyautogui.keyDown('ctrl')
    pyautogui.press('end')
    pyautogui.keyUp('ctrl')
    time.sleep(2)
    pyautogui.moveTo(1450,660) #Come to Start Position
    time.sleep(2)
    pyautogui.mouseDown(button='left') #Press the Left Click Button
    time.sleep(2)
    pyautogui.moveTo(1700,800) #cOme to end Position
    time.sleep(2)
    pyautogui.mouseUp(button='left') #Release the Left Click Button
    time.sleep(2)
    pyautogui.click(x=843,y=437,button='left') #Click on the Name of The Signer
    time.sleep(2)
    pyautogui.click(x=1293,y=792,button='left') #Click on Continue
    time.sleep(2)
    pyautogui.click(x=1338,y=821,button='left') #Click on Sign
    time.sleep(2)
    pth = os.path.dirname(f)
    extension = os.path.splitext(f)[1]
    filename=os.path.splitext(f)[0]
    newfile=os.path.join(pth, filename+'_signed'+extension) #Make a new name for file
    pyautogui.write(newfile) #Give new name to the file
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.write("PASSWORD@1")  #Type in the Password
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.hotkey('alt','f4') #Close the Application of Acrobat


messagebox.showinfo("Output","The Action has been performed. All pdfs has been digitally signed.!")

