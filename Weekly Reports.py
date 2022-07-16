pip install selenium

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import numpy as np
import datetime

options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")

#Open the webdriver
driver = webdriver.Chrome(executable_path=r'C:/Users/edsantil2101/Python/chromedriver.exe')



#Go to the login Page of Pollux IBIS KSA
driver.get('https://itam-saudi-arabia-prod.aws-itam-prod-eu-west-1.nlsn.media/cas/login?service=https://itam-saudi-arabia-prod.aws-itam-prod-eu-west-1.nlsn.media:443/ibisanalytics')

#Get inside the box of "UserName" and the introduce our Username.
id_box = driver.find_element_by_name('username')
id_box.send_keys('edsantil2101')

#Get inside the box of "Password" and the introduce our Password.
password_box = driver.find_element_by_name('password')
password_box.send_keys('Nielsenpassword2.')

#The login button is pressed
login_button = driver.find_element_by_name('submit')
login_button.click()

# Wait until the webpage is loaded or for a maximum of 15 seg
driver.implicitly_wait(15)

#Select the option 'Unknown Channel' Report
#unknown_channel_report="/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[7]/div"
#unknown_channel_report='/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[7]/div/a/span'
panel_report = '/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[1]/ul/li[1]/div/a/span'
report=driver.find_element_by_xpath(panel_report)
report.click()


# Wait until the webpage is loaded or for a maximum of 15 seg
driver.implicitly_wait(15)

#The dates of the period we want are introduced
date_1=driver.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input')
date_1.clear()
date_1.send_keys('2022-06-05')#initial date

date_2=driver.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/input')
date_2.clear()
date_2.send_keys('2022-06-11')#Last date

exe_button=driver.find_element_by_xpath('/html/body/div[8]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/em/button')
exe_button.click()

#########################################
#Wait until the report is generated. Usually it takes from 30 to 35 seconds.
#The time can be changed according how long it takes to load the report in a particular PC running this script
time.sleep(35)

#Changes to the frame where is the report
report_frame = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/div/iframe')
driver.switch_to.frame(report_frame)


#The Button of export is pressed
export_button = driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td[6]/input')
export_button.click()
#Displays the list of the options of format to export and select the option 'Excel'
options = driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[1]/table/tbody/tr[2]/td/select')
options.click()
excel = driver.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[1]/table/tbody/tr[2]/td/select/option[7]")
excel.click()

#Finally press the ok button to make the download
ok_button = driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[2]/div[2]/div/div[1]/input')
ok_button.click()

#give it a break
driver.implicitly_wait(15)

#switch
#selection_frame = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[1]/div/div/iframe') 
#driver.switch_to.frame(selection_frame)
driver.switch_to.default_content()

#Next report
validation_report = '/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[1]/div/a/span'
report2=driver.find_element_by_xpath(validation_report)
report2.click()

# Wait until the webpage is loaded or for a maximum of 15 seg
driver.implicitly_wait(15)

#The dates of the period we want are introduced
date_12=driver.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input')
date_12.clear()
date_12.send_keys('2022-06-05')#initial date


date_22=driver.find_element_by_xpath('/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/input')
date_22.clear()
date_22.send_keys('2022-06-11')#Last date

exe_button2=driver.find_element_by_xpath('/html/body/div[8]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/em/button')
exe_button2.click()

#########################################
#Wait until the report is generated. Usually it takes from 30 to 35 seconds.
#The time can be changed according how long it takes to load the report in a particular PC running this script
time.sleep(35)

#Changes to the frame where is the report
report_frame2 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[3]/div/div/iframe')
driver.switch_to.frame(report_frame2)



#The Button of export is pressed
export_button2 = driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td[6]/input') 
#export_button2.click()
driver.execute_script('arguments[0].click();',export_button2)

#Displays the list of the options of format to export and select the option 'Excel'
options2 = driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[1]/table/tbody/tr[2]/td/select')
#options2.click()
driver.execute_script('arguments[0].click();',options2)

excel2 = driver.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[1]/table/tbody/tr[2]/td/select/option[7]")
excel2.click()
#driver.execute_script('arguments[6].click();',excel2)

#Finally press the ok button to make the download
ok_button2 = driver.find_element_by_xpath('/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[2]/div[2]/div/div[1]/input')
#ok_button2.click()
driver.execute_script('arguments[0].click();',ok_button2)

#Open the downloaded reports
df_PR = pd.read_excel(r'C:\Users\edsantil2101\Downloads\DailyPanelPerformanceSixMonths (1).xls')
df_VS = pd.read_excel(r'C:\Users\edsantil2101\Downloads\DailyValidationSummarySixMonths (1).xls')

#Same process for both inputs
for nameReport, df_WR in [('Panel Report',df_PR), ('Validation Summary',df_VS)]:

    num_rowWR, num_columWR = df_WR.shape

    RowsDrop = list(range(0,6))

    #Drop rows
    if df_WR is df_PR:
        RowsDrop.extend(range(num_rowWR-4,num_rowWR))
    else:
        RowsDrop.extend(range(num_rowWR-2,num_rowWR))

    df_WR1 = df_WR.drop(RowsDrop)
    
    #Get columns indexes
    list1 = df_WR1.columns.tolist()
    list2 = [index for index, value in enumerate(list1)]
    
    #Keep the relevant columns
    if df_WR1[df_WR1.columns[list2[-1]]].isnull().all():
        df_WR1 = df_WR1[df_WR1.columns[[list2[0],list2[-4],list2[-3]]]]
    else:
        df_WR1 = df_WR1[df_WR1.columns[[list2[0],list2[-3],list2[-2]]]]
    
    #Rename columns
    df_WR1.columns = ['KPI','Average (Date Range)','YTD (Date Range)']
    
    nameFile = r'C:\Users\edsantil2101\Python\{}.csv'.format(nameReport)
    
    #Create output
    df_WR1.to_csv(nameFile,na_rep = 'N/A',index = False)
