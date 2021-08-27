#this is a code which will  open a window and that will start to record the screen
#Step 1: Run the code in Pycharm/ Jupyter notebook.  A small window named "Secret Capture" will appear on your taskbar.
#Step 2: It will automatically start recording whatever is there in screen. You have to minimize this secret Capture window. Don't Close it
# Step3: You can continue to change tabs, and everything shown on screen will be recorded

#Note that any kind of sound will NOT be recorded

#Step 4: To stop recording , go to that window "Secret Capture" and press "q"

#Step 5: The recording will be saved in the same folder where your script file is stored.


import datetime

from PIL import ImageGrab
import numpy as np
import cv2
import ctypes

user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
width = user32.GetSystemMetrics(0)
height = user32.GetSystemMetrics(1)
time_stamp = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
file_name = f'{time_stamp}.mp4'
fourcc = cv2.VideoWriter_fourcc('m', 'p', '4', 'v')
captured_video = cv2.VideoWriter(file_name, fourcc, 20.0, (width, height))



while True:
    img = ImageGrab.grab(bbox=(0, 0, width, height))
    img_np = np.array(img)
    img_final = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)

    cv2.imshow('Secret Capture', img_final)



    captured_video.write(img_final)
    if cv2.waitKey(10) == ord('q'):
        break