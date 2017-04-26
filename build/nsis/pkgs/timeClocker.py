#Clocks time in ExponentHR

import re
import getpass
from selenium import webdriver
from selenium.webdriver.support.ui import Select

#points to driver
browser = webdriver.Chrome('C:\Program Files\Time Clocker\chromedriver.exe')

def getCreds():
    #Ask for username & pw
    print("ExponentHR Username:")
    username = input()
    password = getpass.getpass(prompt='ExponentHR Password: (This is a secure input, no characters will appear)\n', stream = None)
    #ensure they're not blank
    if '' in (username,password):
        print("Please fill in all fields.")
        getCreds()
    #assign to tuple
    credsTup = (username,password)
    return(credsTup)

def login(credsTup):
    #Navigates to login page
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
        print('Login incorrect, try again')
        login(getCreds())
        return

def getSecAns():
    #Print security Question and ask for security answer
    #check we're on the secondary auth page and get the question
    pageTitle = str(browser.title)
    if pageTitle == 'Secondary Authentication Required':
        secQuestionElem = browser.find_element_by_id('Hidden2')
        secQuestion = secQuestionElem.get_attribute("value")
    else:
        return
    print("Security Question:")
    print(secQuestion)
    secAnswer = getpass.getpass(prompt='Security Answer:  (This is a secure input, no characters will appear)\n', stream = None)
    #ensure it's not blank
    if secAnswer == '':
        print("Please provide an answer.")
        getSecAns()

    #security field answer & submit
    secField = browser.find_element_by_id('HintA')
    secField.send_keys(secAnswer)

    #submit answer
    answerSubmit = browser.find_element_by_name('BSubmit')
    answerSubmit.click()

    
    #check that we succeeded and got off the secondary auth page to move on
    pageTitle = str(browser.title)
    if pageTitle == 'Secondary Authentication Required':
        print("Security answer incorrect, try again")
        getSecAns()
        return

def main():

    #Navigates to readme
    browser.get('C:\\Program Files\\Time Clocker\\readme.html')

    login(getCreds())
    getSecAns()

    #navigate to timeclock page
    browser.get('https://www.exponenthr.com/service/EmpTimeClock.asp');

    #get the text of the element containing already entered holiday dates
    holiday = browser.find_element_by_id('TableScript1')
    holiText = holiday.text

    #use regex to search the text string for all clocked dates in the xx/xx/xxxx format
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


    #create a list of matching items between the list of dates clocked, and dates in the dropdown
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
        for i in range(y+1,y+6):        
            #jump to the day we are on and print it out
            dateDropDown = Select(browser.find_element_by_name('ThePayDateINOUTTwo'))
            dateDropDown.select_by_index(i)
            print(textList[i])

            #continue if we have a matching date index, effectively skipping a day
            if i in dateIndexes:
                print("Already Clocked or PTO/Holiday, Skipping!\n")
                continue

            #jump to 9AM in time dropdown & submit
            timeDropDown = Select(browser.find_element_by_name('ThePayTimeINOUT'))
            timeDropDown.select_by_value("9:00 AM")
            punchButton = browser.find_element_by_name('B1')
            punchButton.click()

            #jump to 5PM in time dropdown & submit
            timeDropDown = Select(browser.find_element_by_name('ThePayTimeINOUT'))
            timeDropDown.select_by_value("5:00 PM")
            punchButton = browser.find_element_by_name('B1')
            punchButton.click()
            print("Clocked!\n")
        x = x + 1

    print("DONE! You may now close this window.")
