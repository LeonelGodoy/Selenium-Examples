
import os
import pandas as pd
from datetime import datetime
'''
This script will log into EOS and download the 'PSCOGARP00020V04' and 'PSCOGARP00020V01' reports.
Then the script will move the files to the following location: 'C:/Users/user/Documents/EOS Reports/Exports'.
The script should be scheduled to run on the 2nd of each month. Removed website and username/password for security.
'''
current_time = datetime.now().strftime('%m-%d-%Y')
search_date = current_time = datetime.now().strftime('%m/1/%Y')

from selenium import webdriver
# using Chrome to access web
from selenium.webdriver.chrome.options import Options
# loads setings to not have to enter my password
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=C:\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2\\")
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
driver = webdriver.Chrome('C:/Users/user/Downloads/chromedriver.exe',options=options)
# open the website
driver.get('website')

# travels the EOS webpage
driver.switch_to.frame("rbsmain")
driver.find_element_by_name("username").send_keys("user")
driver.find_element_by_name("password").send_keys("pass")
driver.find_element_by_xpath("//input[@onclick='enc(document.logonForm)']").click()
driver.switch_to.default_content()
driver.switch_to.frame("rbstreeframe")
driver.switch_to.frame("rbstreeview")
driver.find_element_by_name("sub7").click()
driver.switch_to.default_content()

# PSCOGARP00020V04
driver.switch_to.frame("rbsmain")
driver.find_element_by_name("reporttblname").send_keys("PSCOGARP00020V04")
driver.find_element_by_name("reporttblcreationDatefrom").send_keys(search_date)
driver.find_element_by_name("reporttblcreationDateto").send_keys(search_date)
driver.find_element_by_name("reporttbluserName").send_keys("DMJGH1")
from selenium.webdriver.common.keys import Keys
driver.find_element_by_name("reporttblname").send_keys(Keys.RETURN)
window_before = driver.window_handles[0]
html = driver.page_source
from bs4 import BeautifulSoup
soup = BeautifulSoup(html, "lxml")
# get *all* 'p' tags with the specified class attribute.
p_tags = soup.findAll('a',{'target':'rbspopup'})
for p in p_tags:
    if "export" in p['href']:
        x = p['href']
        break
# clicks the file and exports a custom excel file
driver.find_element_by_xpath("//a[@href='"+x+"']").click()
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
driver.find_element_by_xpath("//a[@class='style-dialog-tab']").click()
driver.find_element_by_xpath("//option[@value='4']").click()
import time
time.sleep(10)
driver.find_element_by_name("submit").click()
time.sleep(10)

# clicks the file and exports a pdf
driver.switch_to.window(driver.window_handles[0])
driver.switch_to.frame("rbsmain")
driver.find_element_by_xpath("//a[@href='"+x+"']").click()
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
driver.find_element_by_xpath("//input[@name='submit']").click()
time.sleep(10)

# PSCOGARP00020V01
driver.switch_to.window(driver.window_handles[0])
driver.switch_to.frame("rbsmain")
driver.find_element_by_name("reporttblname").clear()
driver.find_element_by_name("reporttblname").send_keys("PSCOGARP00020V01")
driver.find_element_by_name("reporttblcreationDatefrom").send_keys(search_date)
driver.find_element_by_name("reporttblcreationDateto").send_keys(search_date)
driver.find_element_by_name("reporttbluserName").clear()
driver.find_element_by_name("reporttblname").send_keys(Keys.RETURN)
window_before = driver.window_handles[0]
html = driver.page_source
soup = BeautifulSoup(html, "lxml")
# get *all* 'p' tags with the specified class attribute.
p_tags = soup.findAll('a',{'target':'rbspopup'})
for p in p_tags:
    if "export" in p['href']:
        x = p['href']
        break
driver.find_element_by_xpath("//a[@href='"+x+"']").click()
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
driver.find_element_by_xpath("//input[@name='submit']").click()
time.sleep(10)

# logs out and closes the webpage
driver.switch_to.window(driver.window_handles[0])
driver.switch_to.frame("rbsmain")
driver.find_element_by_xpath("//a[@href='logout.doi']").click()
driver.close()

current_time = datetime.now().strftime('%m_%d_%Y')
# moves the file to the appropriate location
path = "C:/Users/user/Downloads/PSCOGARP00020V04.xls"
new_path = "C:/Users/user/Documents/EOS Reports/Exports/"+current_time+" - PSCOGARP00020V04.xls"
os.rename(path, new_path)

path = "C:/Users/user/Downloads/PSCOGARP00020V04.pdf"
new_path = "C:/Users/user/Documents/EOS Reports/Exports/"+current_time+" - PSCOGARP00020V04.pdf"
os.rename(path, new_path)

path = "C:/Users/user/Downloads/PSCOGARP00020V01.pdf"
new_path = "C:/Users/user/Documents/EOS Reports/Exports/"+current_time+" - PSCOGARP00020V01.pdf"
os.rename(path, new_path)
