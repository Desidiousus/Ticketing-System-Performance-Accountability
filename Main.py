from selenium import webdriver
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
import time
import datetime
import os

#Stored Day-Month-Year 
curr_date = datetime.datetime.now().strftime('%d-%m-%Y_%A')

#Create Folder - DeptPerf - DATE
directory = r"Z:\Software\0PerformanceLog\DeptPerf - " + str(curr_date)
directory2 = r"E:\DeptPerf - " + str(curr_date)
try:
    # Create target Directory
    os.mkdir(directory)
    print("Directory " , directory ,  " Created ") 
except FileExistsError:
    print("Directory " , directory ,  " already exists")
try:
    # Create target Directory
    os.mkdir(directory2)
    print("Directory2 " , directory2 ,  " Created ") 
except FileExistsError:
    print("Directory2 " , directory2 ,  " already exists")

while True:
    options = Options()
    options.add_argument("--window-size=250,250")
    CHROMEDRIVER_PATH = r"C:\Users\agyger.ASRCFH\Downloads\chromedriver.exe"

    #Select browser you want to open the url with
    driver = webdriver.Chrome(options = options,executable_path=CHROMEDRIVER_PATH) 

    #URL the browser should go to
    driver.get("http://servicedesk.asrc.com/ProcessManager/Portal/Template20_80.aspx?PageID=54b68520-bd2c-44ca-b646-344400f41ba2")
    #Wait to authenticate
    print("wait 20 seconds...")
    time.sleep(20)
    #Click Service Desk Federal 
    SDF = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_0']/*/img")
    #Click Unclassified
    Unclass = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_3']/*/img")
    #Click Unclassified 2nd level
    Unclass2 = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_3_0']/*/img")
    #Click Federal Problems 
    FP = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_2']/*/img")
    
    #Click Service Desk Federal > Unassigned 
    SDFU = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_0_2']/*/img")
    #Click Service Desk Federal > Massey 
    SDFKM = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_0_9']/*/img")
    #Click Service Desk Federal > Caldwell 
    SDFAC = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_0_3']/*/img")
    #Click Service Desk Federal > Gyger 
    SDFAG = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_0_6']/*/img")

    #Click Federal Problems > Massey
    FPM = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_2_7']/*/img")
    #Click Federal Problems > Caldwell
    FPC = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_2_1']/*/img")
    #Click Federal Problems > Gyger
    FPG = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_2_4']/*/img")
    #Click Federal Problems > Unassigned
    FPU = driver.find_elements_by_xpath("//tr[@id='ctl00xWebPartManager1xwp409527971xctl01xgridWebGrid_r_2_6']/*/img")
    
    #Click SD Federal Drop Down
    if SDF: 
        SDF[0].click()
        
    #Click SD Federal - Unassaigned Drop Down
    if SDFU: 
        SDFU[0].click()
        
    #Click Unclassified Drop Down
    if Unclass: 
        Unclass[0].click()
        time.sleep(5)
    #Click Unclassified 2nd level Drop Down
    if Unclass2: 
        Unclass2[0].click()
        
    #Click Federal Problems Drop Down
    if FP: 
        FP[0].click()
        

    #Click Federal Problems > Massey Drop Down
    if FPM: 
        FPM[0].click()
        
    #Click Federal Problems > Caldwell Drop Down
    if FPC: 
        FPC[0].click()
        
    #Click Federal Problems > Gyger Drop Down
    if FPG: 
        FPG[0].click()
        
    #Click Federal Problems > Unsassigned Drop Down
    if FPU: 
        FPU[0].click()
        

    #Click SD Federal - Massey Drop Down
    if SDFKM: 
        SDFKM[0].click()
        
    #Click SD Federal - Caldwell Drop Down
    if SDFAC: 
        SDFAC[0].click()
        
    #Click SD Federal - Gyger Drop Down
    if SDFAG: 
        SDFAG[0].click()

    #Current time stamp
    curr_time = datetime.datetime.now().strftime('%Hh%Mm%Ss')
    
    #Set Window Siz to ScreenShot Size
    driver.set_window_size(1500, 10000)

    #ScreenShot to Flash Drive
    driver.save_screenshot(r"E:\DeptPerf - " +str(curr_date) + r"\TicketQueueScreenShot - "+str(curr_time)+".png")
    #ScreenShot To Network File Share
    driver.save_screenshot(r"Z:\Software\0PerformanceLog\DeptPerf - " +str(curr_date) + r"\TicketQueueScreenShot - "+str(curr_time)+".png")

    #Close Browser
    driver.quit() 

    print("wait 15 minutes...")
    time.sleep(900)