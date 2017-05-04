#timeClocker 1.2.3 - Forms, Days, Lunches, Choice of Start Times

import re
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time
from urllib.parse import urlparse, parse_qs
import datetime

#points to driver, disables password manager, moves to right side of screen
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    'credentials_enable_service': False,
    'profile': {
        'password_manager_enabled': False
    }
})
chrome_options.add_argument("window-position=1000,0")
chrome_options.add_argument("window-size=900,900")
browser = webdriver.Chrome('C:\Program Files\Time Clocker\chromedriver.exe',chrome_options=chrome_options)

def getCreds():
    #open the credentials form
    browser.get('C:\\Program Files\\Time Clocker\\form.html')

    #get the current URL and wait until it has a "?" in it,
    #which indicates the user has pressed submit
    currentUrl = str(browser.current_url)
    while "?" not in currentUrl:
        time.sleep(1)
        currentUrl = str(browser.current_url)

    #once the user has clicked submit, parse the URL and grab the get values
    parsedUrl = urlparse(currentUrl)
    userDataDict = parse_qs(parsedUrl.query) 
    
    #strip the unnecessary characters from the dict values,
    #make it a list
    credsList = []
    for value in userDataDict.values():
        value = str(value)[2:-2]
        credsList.append(value)

    #ensure nothing is blank
    listCount = len(credsList)
    if listCount < 8:
        print("Please fill in all fields!")
        return getCreds()

    #ensure dayin precedes dayout
    if credsList[2] > credsList[3]:
        print("Please select days in the proper order!")
        return getCreds()
    
    timeList = []
    #take clocked times out and convert them to date format
    for times in range(4,8):
        if "M" not in credsList[times]:
            continue
        else:
            timeList.append(datetime.datetime.strptime(credsList[times],'%I:%M %p'))
            
    #ensure pairs are clocked
    if len(timeList) % 2 == 1:
        print("Please ensure you clock properly!")
        return getCreds()

    #ensure that times don't conflict and are actually clockable
    for item in range(0,len(timeList)-1):
        if timeList[item+1] <= timeList[item]:
            print("Times are unclockable! Please ensure your entered times follow each other.")
            return getCreds()

    #make it a tuple containing name, password, dayin, dayout, time in,
    #lunch start, lunch end, time out
    dataTup = tuple(credsList)
    return dataTup

def login(credsTup):
    #navigates to login page
    browser.get('https://www.exponenthr.com/service/EmpLogon.asp')
    #identify login fields and fill with creds
    emailElem = browser.find_element_by_id('USERID')
    emailElem.send_keys(credsTup[0])
    passwordElem = browser.find_element_by_id('PW')
    passwordElem.send_keys(credsTup[1])

    #submit login
    submitButton = browser.find_element_by_name('B1')
    submitButton.click()

    #check we logged in successfully
    pageTitle = str(browser.title)

    #check that we got the pw right, if not restart
    pageTitle = str(browser.title)
    if pageTitle == 'Validate Failed':
        print('Login incorrect, try again!')
        return login(getCreds())
    else:
        print("Login Successful!\n")

def main():
    #make dataTup a global var
    dataTup = getCreds()
    login(dataTup)
    
    #prompt the user to enter their sec answer then
    #wait til they get to the home page
    print("Please answer your security question in the browser.\n")
    currentUrl = str(browser.current_url)
    while "https://www.exponenthr.com/service/EmpWelcome.asp" not in currentUrl:
        time.sleep(1)
        currentUrl = str(browser.current_url)

    #navigate to timeclock page
    browser.get('https://www.exponenthr.com/service/EmpTimeClock.asp');

    #get the text of the element containing already entered holiday dates
    holiday = browser.find_element_by_id('TableScript1')
    holiText = holiday.text

    #use regex to search the text string for all clocked dates
    #in the xx/xx/xxxx format
    dateMatch = re.findall(r"(\d+/\d+/\d+)",holiText)

    #get a list of values in the date dropdown to compare against
    dateDropDown = Select(browser.find_element_by_name('ThePayDateINOUTTwo'))
    thisWeekOptions = dateDropDown.options
    valuesList = []
    for option in thisWeekOptions:
        valuesList.append(option.get_attribute("value"))

    #create the list of date values for informational use later
    textList = []
    for option in thisWeekOptions:
        textList.append(option.text)

    #create a list of matching items between the list of dates clocked,
    #and dates in the dropdown
    matchSet = list(set(dateMatch).intersection(valuesList))

    #find the all indexes for the matching items and sort
    dateIndexes = []
    for item in matchSet:
        dateIndexes.append(valuesList.index(item))
    dateIndexes.sort(key=int)

    #get how many items are in the date drowdown and div by 7 to get # of weeks
    weeks = (len(dateDropDown.options)/7)
    x = 0

    while x < weeks:
        y = x * 7
        for i in range(y+int(dataTup[2]),y+int(dataTup[3])+1):        
            #jump to the day we are on and print it out
            dateDropDown = Select(browser.find_element_by_name('ThePayDateINOUTTwo'))
            dateDropDown.select_by_index(i)
            print(textList[i])

            #continue if we have a matching date index,
            #effectively skipping a day
            if i in dateIndexes:
                print("Already Clocked or PTO/Holiday, Skipping!\n")
                continue
            
            #loop through the times in the tuple and clock them
            for tupID in range(4,8):
                #skip if blank (for no lunch
                if dataTup[tupID] == 'NA':
                    continue
                else:
                    #jump to in time dropdown & submit
                    timeDropDown = Select(browser.find_element_by_name('ThePayTimeINOUT'))
                    timeDropDown.select_by_value(dataTup[tupID])
                    punchButton = browser.find_element_by_name('B1')
                    punchButton.click()
            
            print("Clocked!\n")
        x = x + 1

    print("DONE! You may now close this window.")
