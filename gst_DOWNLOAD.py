from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
# import request
from selenium.webdriver.common.by import By
import os

path="C:\Program Files (x86)\chromedriver.exe"

driver=webdriver.Chrome(path)

driver.get("https://www.gst.gov.in/")
print(driver.title)
time.sleep(5)

driver.find_element_by_xpath("/html/body/div[1]/header/div[2]/div/div/ul/li/a").click()
time.sleep(5)

ele_user=driver.find_element_by_id("username")
ele_userpass=driver.find_element_by_id("user_pass")

print(ele_userpass.is_displayed())
print(ele_userpass.is_enabled())

user=os.environ.get("gst_bihar_normal_user")
passwd=os.environ.get("gst_bihar_normal_pass")
ele_user.send_keys(user)
ele_userpass.send_keys(passwd)

time.sleep(5)

# Click on the login Button
driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div/div/div/div/div/form/div[6]/div/button").click()
time.sleep(5)

print("Login Button Clicked")

year=["2020"]
quarter=["q","qq","qqq","qqqq"]
months=["Apr","May","Jun"]

#click on the Return Dashboard

driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[3]/div/div[1]/button/span").click()
time.sleep(7)

for each in quarter:
    for mon in months:
        time.sleep(5)

        #Selecting the Financial Year of the Return

        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[1]/select").click()

        fin_year=driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[1]/select")
        fin_year.send_keys(year)


        time.sleep(2)

        # try:
        #     qtr = WebDriverWait(driver, 5).until(
        #         EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[2]/select'))).click()
        #     qtr.send_keys(each)
        # except Exception as e:
        #     print(e)


        #Selcting the Quarter of the Return
        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[2]/select").click()

        qtr = driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[2]/select")

        qtr.send_keys(each)

        #Selecting the Month of the Return

        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[3]/select").click()

        month=driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[3]/select")

        month.send_keys(mon)


        #Clicking on the search Button

        search_key=driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[4]/button")

        search_key.send_keys(Keys.RETURN)

        # driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[4]/div[3]/div[2]/div[1]/div/div/div/div[1]/button")


        print("Entered the Return Dashboard")


        time.sleep(6)

        #Clicking on the GSTR3B view Button

        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[4]/div[3]/div[2]/div[1]/div/div/div/div[1]/button").click()

        # gstr3b_view.send_keys(Keys.RETURN)  while running for the June month the div[3] gets converted to div [4] but for may its fine /html/body/div[2]/div[2]/div/div[2]/div[4]/div[4]/div[2]/div[1]/div/div/div/div[1]/button

        print("Clicked the GSTR-3B View Button")


        time.sleep(6)

        # driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/button[2]").click()
        #
        # time.sleep(5)






        try:
            download_but = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/button[2]'))).click()

        except Exception as e:
            print(e)

        print("Clicked the GSTR-3B Download Button")


        driver.find_element_by_xpath("/html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/a").click()
        time.sleep(5)


        print("Clicked the Baack Button")



# gsR-1 vIEW BUTTON >>> /html/body/div[2]/div[2]/div/div[2]/div[4]/div[3]/div[1]/div[1]/div/div/div/div[1]/button

# gsR-3B vIEW BUTTON >>>/html/body/div[2]/div[2]/div/div[2]/div[4]/div[3]/div[2]/div[1]/div/div/div/div[1]/button

# /html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[2]/select >>> This is the x path for the FY 2019-20 and prior

# /html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[3]/select >>> This is the X path for the FY 2020-21 and beyond

# /html/body/div[2]/div[2]/div/div[2]/div[2]/form/div/div[4]/button >>> This is the Xpath for the search Button


# /html/body/div[2]/div[2]/div/div[1]/div[1]/div/ol/li[2]/ng-switch/span/a

# /html/body/div[2]/div[2]/div/div[1]/div[1]/div/ol/li[2]/ng-switch/span/a

# /html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/a
#
# /html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/button[2]
# /html/body/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[9]/div/button[2]


time.sleep(5)




driver.close()