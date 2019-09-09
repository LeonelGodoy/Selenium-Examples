import os
import pandas as pd
from datetime import datetime
from selenium import webdriver
# using Chrome to access web
from selenium.webdriver.chrome.options import Options
'''
Downloads a excel file that is then compared that of the previous export.
'''

current_time = datetime.now().strftime('%m-%d-%Y-%H-%M')
# loads setings to not have to enter my password
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument("--user-data-dir=C:"+
"\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data\\\\Profile 2\\")
#options.add_argument("--disable-extensions")
#options.add_argument("--disable-gpu")
#options.add_argument("--headless") This script cannot be headless....

driver = webdriver.Chrome('C:/Users/user/Downloads/'+
'chromedriver.exe',options=options)

# open the website
driver.get('https://www.snl.com/interactivex/bbsearch.aspx?'+
'activeTabIndex=8&showLeftNavigation=1&printable=1')

driver.find_element_by_css_selector(".SNLToolBarItem").click()

driver.find_element_by_id("1077155699").click()

driver.find_element_by_id("openButton_saveControl").click()

driver.find_element_by_xpath("//div[@onclick='"+
"SaveControl_saveControl.OpenConfirmClick();']").click()

search = driver.find_element_by_name("btnSearchTop")
search.click()

driver.find_element_by_id("exportsMenuItem").click()

driver.find_element_by_xpath("//li[@style='width: 110px; margin-left:"+
" -37px; vertical-align: middle; float: left;'][2]").click()

# sleeps so that the file can be properly downloaded
import time
time.sleep(10)

old_path = "C:/Users/user/Downloads/snlworkbook.xls"
new_path = "C:/Users/user/Documents/S&L Automation/Current Search/S&L - HOM 3 Month.xls"
previous_file = "C:/Users/user/Documents/S&L Automation/Previous Search/S&L - HOM 3 Month.xls"
delete_path = "C:/Users/user/Documents/S&L Automation/Archive/" + current_time + "-S&L - HOM 3 Month.xls"
os.rename(old_path, new_path)

# loads the previous search and the curreent search, filling NA values
df0 = pd.read_excel(new_path, sheet_name='Product Filings Search',header =9)
new_search_data = df0.iloc[:-5,:].drop(columns=['Snippet']).fillna("NA")

df1 = pd.read_excel(previous_file, sheet_name='Product Filings Search',header =9)
previous_search_data = df1.iloc[:-5,:].drop(columns=['Snippet']).fillna("NA")

new_search_data["SERFF Index"] = new_search_data['SERFF Tracking # /State Tracking #']
previous_search_data["SERFF Index"] = previous_search_data['SERFF Tracking # /State Tracking #']

def highlight_cols(x):
    # Copy df to new - original data are not changed
    df = x.copy()
    df.iloc[:,:] = 'color: black'
    # Color the cell green if the value is a new entry
    for n,value in enumerate(x['SERFF Tracking # /State Tracking #']):
        if value in list(previous_search_data['SERFF Tracking # /State Tracking #']):
            df.iloc[n,0] = 'color: black'
        else:
            df.iloc[n,0] = 'background-color: #2ca02c'
    # Colors the rate change depending on if they are greater than or less than 0
    # for n,value in enumerate(x['Rate Change (%)¹']):
    #     if value == "NA":
    #         df.iloc[n,10] = 'color: black'
    #     elif value < 0:
    #         df.iloc[n,10] = 'color: red'
    #     elif value > 0:
    #         df.iloc[n,10] = 'color: green'
    #     else:
    #         df.iloc[n,10] = 'color: black'


    # Highlights the background of a cell if it has been changed since the last search
    data_index_previous = previous_search_data.set_index('SERFF Index')
    data_index_new = new_search_data.set_index('SERFF Index')
    data_index_previous = data_index_previous[data_index_previous.index.isin(list(data_index_new.index))]

    for i in data_index_previous.index:
        for c in data_index_previous:
            if data_index_previous.loc[i,c] == data_index_new.loc[i,c]:
                df.loc[i,c] = df.loc[i,c] + '; background-color: white'
            else:
                df.loc[i,c] = df.loc[i,c] + '; background-color: yellow'
    return(df)

# exports the changelog
new_search_data.set_index('SERFF Index').style.apply(
highlight_cols, axis=None).to_excel("C:/Users/user/Documents/S&L Automation/HOM/" +
 current_time + " HOM changelog.xlsx", index=False)

os.rename(previous_file, delete_path)
os.rename(new_path, previous_file)

driver.close()
