# ==================================== #
# FORM_FETCH is a web scraping utility
# This file includes function definitions for searching HTML and entering text into a corresponding text field or clicking corresponding buttons
# The current script will download the 'desired_form's form excel sheet from a Microsoft account
# File Paths Start With Script Directory
# ==================================== #

import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options


#Modifiable Data
desired_form            = 'FORM NAME'
username                = 'USERNAME'
pswrd                   = 'PASSWORD'
firefox_driver_path     = 'geckodriver'
download_path           = ''
url                     = "https://login.microsoftonline.com/common/oauth2/authorize?response_mode=form_post&response_type=id_token&scope=openid&msafed=0&nonce=2f63076a-6394-4359-be60-5525b99fdb9f.636519878208739339&state=https%3A%2F%2Fforms.office.com%2FPages%2FDesignPage.aspx%3Forigin%3Dshell&client_id=c9a559d2-7aab-4f13-a6ed-e7e9c52aec87&redirect_uri=https:%2f%2fforms.office.com%2fauth%2fsignin"
delay_sleep             = .5




#Redirect Paths To Relative Script Path
download_path = os.path.dirname(os.path.realpath(__file__)) + '\\' + download_path
firefox_driver_path = os.path.dirname(os.path.realpath(__file__)) + '\\' + firefox_driver_path

#FireFox Options
options = Options()
options.add_argument("--headless")

#FireFox Preferences
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", download_path)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', "text/plain, application/vnd.ms-excel, text/csv, application/csv, text/comma-separated-values, application/download, application/octet-stream, binary/octet-stream, application/binary, application/x-unknown")

#Open Browser
driver = webdriver.Firefox(executable_path=firefox_driver_path, firefox_profile=profile, firefox_options=options)
driver.get(url)
print("Initialized")

def EnterText(element_id, value):
    try:
        inputElement = driver.find_element_by_id(element_id);
        inputElement.send_keys(value)
    except:
        sleep(delay_sleep)
        EnterText(element_id, value)

def ClickButton(element_id, id_search = True):
    try:
        if id_search :
            inputElement = driver.find_element_by_id(element_id).click();
        else :
            inputElement = driver.find_element_by_xpath(element_id).click();
    except:
        sleep(delay_sleep)
        ClickButton(element_id, id_search)


#Script
EnterText('i0116', username)                                #Username
ClickButton('idSIButton9')                                  #Next
EnterText('i0118', pswrd)                                   #Password
ClickButton('idSIButton9')                                  #Sign In
ClickButton('idBtn_Back')                                   #Stay Signed In? No
ClickButton('//*[@title="' + desired_form + '"]', False)    #Desired Form
ClickButton('//*[@aria-label="Responses"]', False)          #Responses
ClickButton('//*[@title="Open in Excel"]', False)           #Download
print("Downloading...")
sleep(5)                                                    #Give Time To Download
driver.close()                                              #Close
print("Done");
