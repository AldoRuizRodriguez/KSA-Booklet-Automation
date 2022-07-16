from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time

def login():#login To the Ibis Webpage
    ibis_webpage='https://itam-saudi-arabia-prod.aws-itam-prod-eu-west-1.nlsn.media/cas/login?service=https://itam-saudi-arabia-prod.aws-itam-prod-eu-west-1.nlsn.media:443/ibisanalytics'

    user_name='alruiz2201'
    password='Ald0.Ru1z22!'

    driver.get(ibis_webpage)

    user_name_box = driver.find_element_by_name('username')
    user_name_box.send_keys(user_name)

    password_box = driver.find_element_by_name('password')
    password_box.send_keys(password)

    #The login button is pressed
    login_button = driver.find_element_by_name('submit')
    login_button.click()

    driver.implicitly_wait(15)

def export_func():#Export to 'xls' each report
    #The Button of export is pressed
    export_button = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td[6]/input")
    export_button.click()

    #Displays the list of the options of format to export and select the option 'Excel'
    options = driver.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[1]/table/tbody/tr[2]/td/select")
    options.click()
    excel = driver.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[1]/table/tbody/tr[2]/td/select/option[7]")
    excel.click()

    #Finally press the ok button to make the download
    ok_button = driver.find_element_by_xpath("/html/body/table/tbody/tr[2]/td[1]/div[4]/div[2]/div/div[2]/div[2]/div/div[1]/input")
    ok_button.click()

def unknown_chanel():
    unknown_channel_report="/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[7]/div"
    report=driver.find_element_by_xpath(unknown_channel_report)
    report.click()
        # Wait until the webpage is loaded or for a maximum of 15 seg
    driver.implicitly_wait(15)

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    exe_button=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    #Changes to the frame where is the report
    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    #Waits until the report is generated
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[2]/tbody/tr/td')))

    export_func()

def nil_viewers_ind():

    driver.switch_to.default_content()

    intab_report = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[13]/div")
    intab_report.click()

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    mode=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[5]/div[1]/div[1]/input[2]")
    mode.clear()
    mode.send_keys('Consecutive')

    exe_button=driver.find_element_by_xpath("/html/body/div[8]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    #Changes to the frame where is the report
    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[3]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    Elem = WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[2]/tbody/tr/td')))
    export_func()

def nil_viewers_hh():
    driver.switch_to.default_content()

    intab_report = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[11]/div")
    intab_report.click()
    driver.implicitly_wait(100)
    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[10]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[10]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    mode=driver.find_element_by_xpath("/html/body/div[10]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[5]/div[1]/div/input[2]")
    mode.clear()
    mode.send_keys('Consecutive')

    exe_button=driver.find_element_by_xpath("/html/body/div[10]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[4]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    Elem = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[2]/tbody/tr/td')))

    export_func()

def off_directory():

    driver.switch_to.default_content()

    intab_report = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[1]/ul/li[4]/div")
    intab_report.click()
    driver.implicitly_wait(60)

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[3]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    exe_button=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[5]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    Elem = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[2]/tbody/tr/td')))

    export_func()

def uncovered_viewing():
    
    driver.switch_to.default_content()

    #Select the option 'Uncovered Viewing Report' Report
    uncovered_viewing_report="/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[9]/div"
    report=driver.find_element_by_xpath(uncovered_viewing_report)
    report.click()

    # Wait until the webpage is loaded or for a maximum of 15 seg
    driver.implicitly_wait(15)

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    exe_button=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    #Changes to the frame where is the report
    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[6]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    Elem = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[2]/tbody/tr/td')))

    export_func()

def rejected_hh():

    driver.switch_to.default_content()

    #Select the option 'Uncovered Viewing Report' Report
    continiuous_rejected_hh="/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[3]/div"
    report=driver.find_element_by_xpath(continiuous_rejected_hh)
    report.click()

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    exe_button=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[7]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    Elem = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div/div/table/tbody/tr[2]/td/table[3]/tbody/tr/td')))
    
    export_func()

def rejected_individuals():
    driver.switch_to.default_content()

    rejected_individuals = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[6]/div")
    rejected_individuals.click()
    driver.implicitly_wait(60)

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#Last date

    accum_days=driver.find_element_by_xpath('/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[3]/div[1]/input')
    accum_days.clear()
    accum_days.send_keys('28')#Last date

    exe_button=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[8]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    time.sleep(30)

    export_func()

def periodic_exceptions():
    driver.switch_to.default_content()

    periodic_exceptions="/html/body/div[1]/div/div/div/div/div/div[3]/div[2]/div/div[2]/div[2]/div[2]/ul/div/li[1]/ul/li[3]/ul/li[23]/div"
    periodic_exceptions=driver.find_element_by_xpath(periodic_exceptions)
    periodic_exceptions.click()

    #The dates of the period we want are introduced
    date_1=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input")
    date_1.clear()
    date_1.send_keys('2022-05-01')#initial date

    date_2=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[2]/div[1]/div[1]/input")
    date_2.clear()
    date_2.send_keys('2022-05-28')#Last date

    mode=driver.find_element_by_xpath("/html/body/div[11]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[6]/div[1]/div/input[2]")
    mode.clear()
    mode.send_keys('Not Consecutive')#Last date
    #driver.find_element_by_xpath("/html/body/div[13]/div[1]/div").click()

    excep=driver.find_element_by_xpath("/html/body/div[10]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[8]/div[1]/div/input[2]")
    excep.clear()
    excep.send_keys('NOT POLLED')

    driver.find_element_by_xpath('/html/body/div[10]/div[2]/div[1]/div/div/div/div/div[2]/div[1]/div/div/div/form/fieldset/div[2]/div[1]/div/div/div/div[1]/div[1]/div[1]/input').click()

    time.sleep(3)
    exe_button=driver.find_element_by_xpath("/html/body/div[10]/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]")
    exe_button.click()

    report_frame = driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/div/div[2]/div[2]/div/div[9]/div/div/iframe")
    driver.switch_to.frame(report_frame)

    time.sleep(30)

    export_func()



executable_path=r'C:/Users/alruiz2201/Documents/chromium/chromedriver.exe'
driver = webdriver.Chrome(executable_path)

login()

#The Reports are generated and exported
unknown_chanel()
nil_viewers_ind()
nil_viewers_hh()
off_directory()
uncovered_viewing()
rejected_hh()
rejected_individuals()
periodic_exceptions()

#Waits 60 sec and closes the WebDriver Navigator
time.sleep(60)
driver.quit()
