from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time


#Initialize the webdriver and goes to the website
executable_path=r'C:/Users/' #Set the route of your selenium webrowser
driver = webdriver.Chrome(executable_path)
spotfire_webpage='http://opintel-itam-spotfire-prod.aws-opintel-us-east-1.nlsn.media/spotfire/wp/analysis?file=/ITAM/Visualizations/KSA/SaudiArabiaProd_OTT_Collections&waid=OeEnzWyd-0aQ5fvQ3lZ9--1515246242AAuJ&wavid=0'    
driver.get(spotfire_webpage)
driver.implicitly_wait(20)

#The page is redirected to OKTA Authentication and our mail is set in the requested field
username_box=driver.find_element_by_id('idp-discovery-username')
username_box.send_keys('') #introduce your mail
next_button=driver.find_element_by_id('idp-discovery-submit').click()

#The program waits until the the Authentication progress is done and the Spotfire webpage is loaded
WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[2]/div/div[5]/div[2]/div[1]/div')))
time.sleep(3)
#Clicks the 'Mapped devices' button report
mapped_dev_button=driver.find_element_by_xpath('/html/body/div/div[2]/div/div[1]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[32]')
driver.execute_script("arguments[0].click();", mapped_dev_button)

time.sleep(1)

#Locates and click the File button
driver.find_element_by_xpath('/html/body/div/div[2]/div/div[5]/div[1]/div[3]/div[1]').click()

#Makes the action of mouse-hover for the option 'export' 
a= ActionChains(driver)
exp= driver.find_element_by_xpath("/html/body/div[2]/div[1]/div[5]")
a.move_to_element(exp).perform()

#Locates and click the option 'Table' to download the data.
driver.find_element_by_xpath('/html/body/div[3]/div[1]/div[3]').click()

#ends the program
time.sleep(5)
driver.quit()

