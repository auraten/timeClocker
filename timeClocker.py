#timeClocker 2.0.4 - UltiPro Adaptation

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
    browser.get('https://nw12.ultipro.com/login.aspx')
    #identify login fields and fill with creds
    emailElem = browser.find_element_by_id('ctl00_Content_Login1_UserName')
    emailElem.send_keys(credsTup[0])
    passwordElem = browser.find_element_by_id('ctl00_Content_Login1_Password')
    passwordElem.send_keys(credsTup[1])

    #submit login
    submitButton = browser.find_element_by_id('ctl00_Content_Login1_LoginButton')
    submitButton.click()
    
    #check that we got the pw right, if not restart
    currentUrl = str(browser.current_url)
    if currentUrl == 'https://nw12.ultipro.com/login.aspx':
        print('Please try your credentials again in the browser, or fill in your security questions.')
    while currentUrl == 'https://nw12.ultipro.com/login.aspx':
        time.sleep(1)
        currentUrl = str(browser.current_url)
    
def main():
    #make dataTup a global var
    dataTup = getCreds()
    login(dataTup)

    print('Login Successful!\n')
    
    #navigate to timeclock page
    menuButton = browser.find_element_by_name('menuButton')
    menuButton.click()
    #this method of navigation replaces the unreliable xpath method I was using
    #and replaces it with more permanent data fields
    #we get a list of all clases, then find the one with the correct data field
    myselfButton = browser.find_elements_by_class_name('menuTopHeader')
    for i in myselfButton:
        if i.get_attribute('data-uitoggle') == 'menu_myself':
            myselfButton = i
    myselfButton.click()
    timeMGMT = browser.find_elements_by_class_name('menuItem')
    for i in timeMGMT:
        if i.get_attribute('data-id') == '2148':
            timeMGMT = i
    timeMGMT.click()

    #drop down into the proper frame
    frame = browser.find_element_by_id('ContentFrame')
    browser.switch_to.frame(frame)
    frame = browser.find_element_by_id('ctl00_Content_UTMFrame')
    browser.switch_to.frame(frame)
    frame = browser.find_element_by_id('MainContentFrame')
    browser.switch_to.frame(frame)
    frame = browser.find_element_by_id('main_frame')
    browser.switch_to.frame(frame)

    #check that we are in the current pay period
    #if not, instruct to stop and submit previous sheet
    payPeriodDD = Select(browser.find_element_by_id('COMBO1_PAYCYCLE_dlDateSelection'))
    if payPeriodDD.first_selected_option.text != 'Current Pay Period':
        print('Please submit your last pay period in UltiPro, then re-run this program.\nThis window will close in 15 seconds.')
        time.sleep(15)
        quit()

    #check that the status of current period is open
    #if not, instruct to stop and troubleshoot
    payPeriodStatus = browser.find_element_by_id('tsStatus')
    if payPeriodStatus.text != 'OPEN':
        print('Please ensure the current pay period is in the open status, then re-run this program.\nThis window will close in 15 seconds.')
        time.sleep(15)
        quit()
    
    #grab the string that specifies the date range
    dateString = browser.find_element_by_class_name('timesheetHeaderText')
    dateText = dateString.text

    #split it
    dates = dateText.split(' to ')

    #convert it to datetime objects, do math to determine if 2 or 3 week period
    startStr = dates[0]
    endStr = dates[1]
    start = datetime.datetime.strptime(startStr, 'My Timesheet for %B %d, %Y')
    end = datetime.datetime.strptime(endStr, '%B %d, %Y')
    w = end - start
    if w.days < 14:
        weeks = 2
    else:
        weeks = 3

    #counter for rows
    z = 0
    
    #fetch list for already clocked dates
    clockedDates = []
    rowStr = 'gdvTS_rw_' + str(z) + '_cl_1'
    dateField = browser.find_element_by_id(rowStr)
    dateText = dateField.text

    #loop through filled rows and append those dates to a list
    while dateText != '':
        clockedDates.append(dateText)
        z = z + 1
        rowStr = 'gdvTS_rw_' + str(z) + '_cl_1'
        dateField = browser.find_element_by_id(rowStr)
        dateText = dateField.text

    #click the date field
    emptyRow = len(clockedDates) + 1

    dateField = browser.find_element_by_id('gdvTS_rw_' + str(emptyRow) + '_cl_1')
    browser.execute_script("return arguments[0].scrollIntoView();", dateField)
    dateField.click()

    #get a list of ALL values in the date dropdown to compare against
    dateDropDown = Select(browser.find_element_by_id('gdvTS_rw_' + str(emptyRow) + '_TPDATE_slc'))
    periodDates = dateDropDown.options

    #convert to text list
    textList = []
    for option in periodDates:
        textList.append(option.text)

    #instantiate list of valid dates
    validDates = []

    #counter for weeks
    x = 0
    
    #run through the users chosen dates to compare clocked dates against the list of all dates to create a list of valid dates
    while x < weeks:
        y = x * 7
        x = x + 1
        for i in range(y+int(dataTup[2])+1,y+int(dataTup[3])+2):
            if textList[i] not in clockedDates:
                validDates.append(textList[i])

    #loop to ensure enough rows exist
    rowCount = len(validDates)
    for _ in range(rowCount):
        #click the add button
        addButton = browser.find_element_by_id('ImgAddRow')
        browser.execute_script("return arguments[0].scrollIntoView();", addButton)
        addButton.click()
    
    #find the index numbers of those valid dates in the list of ALL dates
    indexList = []
    for i in validDates:
        indexList.append(textList.index(i))

##    #find index #s of clocked dates and print out that they are already clocked
##    clockedIndexList = []
##    for i in clockedDates:
##        clockedIndexList.append(textList.index(i))
##
##    for i in clockedIndexList:
##        print(str(textList[clockedIndexList[i]]) + ' has already been clocked!\n')

    #clock based on only the valid dates
    #z counter is already iterated properly to avoid row collisions
    for i in indexList:   
        #print the date we're working on
        print(textList[i])

        #insert the row counter into the string to interate rows
        rowStr = 'gdvTS_rw_' + str(z) + '_cl_1'
        dateField = browser.find_element_by_id(rowStr)
        
        #click the date field
        browser.execute_script("return arguments[0].scrollIntoView();", dateField)
        dateField.click()

        #insert the counter into the string to iterate dates in the drop down
        dropDownStr = 'gdvTS_rw_' + str(z) + '_TPDATE_slc'
        dateDropDown = Select(browser.find_element_by_id(dropDownStr))
        dateDropDown.select_by_index(i)

        #open the pay code and select work
        payStr = 'gdvTS_rw_' + str(z) + '_cl_2'
        payField = browser.find_element_by_id(payStr)
        payField.click()
        payDropDownStr = 'gdvTS_rw_' + str(z) + '_NPAYCODE_slc'
        payDropDown = Select(browser.find_element_by_id(payDropDownStr))
        payDropDown.select_by_index(1)

        #insert the times
        timeStr = 'gdvTS_rw_' + str(z) + '_cl_3'
        timeBox = browser.find_element_by_id(timeStr)
        timeBox.click()
        timeFieldStr = 'gdvTS_rw_' + str(z) + '_TIN_txt'
        timeField = browser.find_element_by_id(timeFieldStr)
        timeField.send_keys(dataTup[4])
        
        if dataTup[5] != 'NA':
            #insert the lunch start time
            timeStr = 'gdvTS_rw_' + str(z) + '_cl_4'
            timeBox = browser.find_element_by_id(timeStr)
            timeBox.click()
            timeFieldStr = 'gdvTS_rw_' + str(z) + '_TOUT_txt'
            timeField = browser.find_element_by_id(timeFieldStr)
            timeField.send_keys(dataTup[5])

            #iterate the row counter
            z = z + 1

            #click the add button
            addButton = browser.find_element_by_id('ImgAddRow')
            browser.execute_script("return arguments[0].scrollIntoView();", addButton)
            addButton.click()

            #insert the row counter into the string to interate rows
            rowStr = 'gdvTS_rw_' + str(z) + '_cl_1'
            
            #click the date field
            dateField = browser.find_element_by_id(rowStr)
            browser.execute_script("return arguments[0].scrollIntoView();", dateField)
            dateField.click()

            #insert the counter into the string to iterate dates in the drop down
            dropDownStr = 'gdvTS_rw_' + str(z) + '_TPDATE_slc'
            dateDropDown = Select(browser.find_element_by_id(dropDownStr))
            dateDropDown.select_by_index(i)

            #open the pay code and select work
            payStr = 'gdvTS_rw_' + str(z) + '_cl_2'
            payField = browser.find_element_by_id(payStr)
            payField.click()
            payDropDownStr = 'gdvTS_rw_' + str(z) + '_NPAYCODE_slc'
            payDropDown = Select(browser.find_element_by_id(payDropDownStr))
            payDropDown.select_by_index(1)

            #insert the resume work time
            timeStr = 'gdvTS_rw_' + str(z) + '_cl_3'
            timeBox = browser.find_element_by_id(timeStr)
            timeBox.click()
            timeFieldStr = 'gdvTS_rw_' + str(z) + '_TIN_txt'
            timeField = browser.find_element_by_id(timeFieldStr)
            timeField.send_keys(dataTup[6])

        #insert the end of day time
        timeStr = 'gdvTS_rw_' + str(z) + '_cl_4'
        timeBox = browser.find_element_by_id(timeStr)
        timeBox.click()
        timeFieldStr = 'gdvTS_rw_' + str(z) + '_TOUT_txt'
        timeField = browser.find_element_by_id(timeFieldStr)
        timeField.send_keys(dataTup[7])

        print('CLOCKED!\n')

        #iterate the row counter
        z = z + 1

    submitButton = browser.find_element_by_id('btn_Apply')
    browser.execute_script("return arguments[0].scrollIntoView();", submitButton)
    submitButton.click()

    print("DONE! Your requested times have been clocked and saved.\nPlease review them for accuracy, then close this window.")
