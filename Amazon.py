import time
import sys
import os
import pandas as pd
from datetime import datetime
import pyautogui
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
# Using Chrome to access web
from selenium.webdriver.chrome.options import Options
'''
Checks the Amazon website for job posting in the Macon Area. If a posting is found it will open a webpage.
'''
while True:
    time.sleep(60)
    # Loads setings to not have to enter my password
    current_time = datetime.now().strftime('%m-%d-%Y-%H-%M-%S')
    print("Running Script at " + current_time)
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\")
    options.add_argument("--headless")
    driver = webdriver.Chrome('C:/Users/User/Documents/Python/Chrome Automation/chromedriver',options=options)
    # Open the website
    driver.get('https://amazon.force.com/Index')
    time.sleep(2)
    driver.find_element_by_xpath("//a[@class='accordion-toggle collapsed']").click()
    time.sleep(2)
    driver.find_element_by_xpath("//input[@id='j_id0:portId:j_id56:chkFC']").click()
    time.sleep(2)
    driver.find_element_by_xpath("//option[@value='US']").click()
    time.sleep(2)
    driver.find_element_by_xpath("//input[@class='form-control input-sm']").send_keys("Macon")
    time.sleep(2)
    driver.find_element_by_xpath("//input[@class='btn btn-gold']").click()
    time.sleep(2)
    try:
        time.sleep(2)
        driver.find_element_by_xpath("//span[@id='j_id0:portId:j_id56:j_id75']").text
        _text = driver.find_element_by_xpath("//span[@id='j_id0:portId:j_id56:j_id75']").text
        print(_text)
        driver.close()
        print("No Job Posting!")
        if "No jobs found" == _text:
            print("Ending Script...")
    except NoSuchElementException:
        driver.close()
        options.add_argument("--user-data-dir=C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\")
        driver.get('https://amazon.force.com/Index')
        time.sleep(2)
        driver.find_element_by_xpath("//a[@class='accordion-toggle collapsed']").click()
        time.sleep(2)
        driver.find_element_by_xpath("//input[@id='j_id0:portId:j_id56:chkFC']").click()
        time.sleep(2)
        driver.find_element_by_xpath("//option[@value='US']").click()
        time.sleep(2)
        driver.find_element_by_xpath("//input[@class='form-control input-sm']").send_keys("Macon")
        time.sleep(2)
        driver.find_element_by_xpath("//input[@class='btn btn-gold']").click()
        time.sleep(2)
        break
quit()
