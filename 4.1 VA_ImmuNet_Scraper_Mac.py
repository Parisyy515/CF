
# Filename: VA_ImmuNet_Scraper_Mac.py
# Author: Zheng Guo
# Date: 12-07-2020
# Purpose: Scraping member's immunization registration information from the VA Immunet site based on a given list of members not found from MD Immunet.
# Class list: - Person (measYr, memberId, memberIdSkey, fname, lname, lnameSuffix, dob, gender, stateRes, meas)

# Functions:
# - is_date(string)
# - immunte(Fname, Lname, DOB, Gender, user, pw)

# User Input:
# - First input : HEDIS_MD_Immun_Records_Not_Found.xlsx
#             xlsx file contains the following information in order below:
#             MEMB_YR,MEMB_LIFE_ID_SKEY,MEMB_LIFE_ID,MEMB_FRST_NM,MEMB_LAST_NM,MEMB_SUFFIX,DOB,GNDR,RSDNC_STATE
# - Second input : User name
# - Third input: Password

# Output:
# - First output: HEDIS_VA_Immun_Records_Found_YYYY_MM_DD.csv
# - Second output: HEDIS_VA_Immun_Records_Not_Found_YYYY_MM_DD.csv
##############################################################################


# Imports
import csv
import datetime
import os
import os.path
import time

import pandas as pd
from dateutil.parser import parse
from pandas import DataFrame
from selenium import webdriver
from selenium.common.exceptions import (UnexpectedAlertPresentException,
                                        WebDriverException)
from selenium.webdriver.support.select import Select

##############################################################################
# Classes


class Person(object):
    def __init__(self, measYr, memberId, memberIdSkey, fname, lname, lnameSuffix, dob, gender, stateRes, meas):
        self.measYr = measYr
        self.memberId = memberId
        self.memberIdSkey = memberIdSkey
        self.fname = fname
        self.lname = lname
        self.lnameSuffix = lnameSuffix
        self.dob = dob
        self.gender = gender
        self.stateRes = stateRes
        self.meas = meas

    def getMeasYr(self):
        return self.measYr

    def getMemberIdSkey(self):
        return self.memberIdSkey

    def getMemberId(self):
        return self.memberId

    def getFirstName(self):
        return self.fname

    def getLastName(self):
        return self.lname

    def getLastNameSuffix(self):
        return self.lnameSuffix

    def getDateOfBirth(self):
        return self.dob

    def getGender(self):
        return self.gender

    def getStateRes(self):
        return self.stateRes

    def getMeas(self):
        return self.meas

###############################################################################
# Function


def is_date(string, fuzzy=False):
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False


def immunte(Fname, Lname, DOB, Gender, driver):

    # work on last name
    lastname = driver.find_element_by_id("txtLastName")
    lastname.clear()
    lastname.send_keys(Lname)

    # work on first name
    firstname = driver.find_element_by_id("txtFirstName")
    firstname.clear()
    firstname.send_keys(Fname)

    # work on birth date
    birthdate = driver.find_element_by_name("txtBirthDate")
    birthdate.clear()
    birthdate.send_keys(DOB)

    # work on gender selection button
    if Gender == 'F':
        driver.find_element_by_xpath("//input[@value='F']").click()
    elif Gender == 'M':
        driver.find_element_by_xpath("//input[@value='M']").click()
    else:
        driver.find_element_by_xpath("//input[@value='N']").click()

    time.sleep(0.25)
    # work on search button

    driver.find_element_by_name("cmdFindClient").click()

    # two scenarios could emerge as a search result: 1, no patient found 2, the patient found
    header = driver.find_element_by_css_selector('p.large').text

    if "Client Search Criteria" in header:
        al = []
        # work on returning to home page
        driver.find_element_by_xpath(
            "//*[@id='xMenu1a']/font/a/font").click()
        header = ''

    elif "Access Restricted" in header:
        al = []
        # work on returning to home page
        driver.find_element_by_xpath(
            "//*[@id='xMenu1a']/font/a/font").click()
        header = ''

    elif "Client Information" in header:

        even = driver.find_elements_by_class_name("evenRow")
        odd = driver.find_elements_by_class_name("oddRow")
        o = []
        e = []
        o1 = []
        e1 = []

        for value in odd:
            o1.append(value.text)
        for value in even:
            e1.append(value.text)

        o = list(filter(lambda x: "/" in x, o1))
        e = list(filter(lambda x: "/" in x, e1))

        length = min(len(o), len(e))
        i = 0
        al = []

        # merge odd and even row together and remove the row marked with complete
        while i < length:
            al.append(e[i])
            al.append(o[i])
            i = i+1

        # parse each row of information with a comma, add group name for row that are without one
        for x in range(len(al)):
            if is_date(al[x][1:10]):
                al[x] = al[x].replace(' ', ',')
                al[x] = al[x].replace(',of,', ' of ')
                al[x] = group + ',' + al[x][2:]

            else:
                al[x] = al[x].replace(' ', ',')
                al[x] = al[x].replace(',of,', ' of ')
                g = al[x].split(',', 1)
                group = g[0]

        # wait for result to load into al
        time.sleep(0.25)
        header = ''
        # work on returning to home page
        driver.find_element_by_xpath(
            "/html/body/table/tbody/tr/td[1]/div/font/a/font").click()
    return al


def main():
    # Welcome message and input info
    print('\nThis is the web scraper for the Virginia Immunization Record Website.')
    print('You will be prompted to type in a file name and username/password.')
    print('If you need to exit the script and stop its process press \'CTRL\' + \'C\'.')
    file = input("\nEnter file name: ")
    org = 'HP02'
    user = input("\nEnter VAImmnet username: ")
    pw = input("\nEnter VAImmnet password: ")
    date = str(datetime.date.today())

    # output file
    fileOutputName = 'HEDIS_VA_Immun_Records_Found_' + \
        date.replace('-', '_') + '.csv'
    fileOutputNameNotFound = 'HEDIS_VA_Immun_Records_Not_Found_' + \
        date.replace('-', '_') + '.csv'

    fileOutput = open(fileOutputName, 'w')
    fileOutputNotFound = open(fileOutputNameNotFound, 'w')

    fileOutput.write('MEAS_YR,MEMB_LIFE_ID_SKEY,MEMB_LIFE_ID,MEMB_FRST_NM,MEMB_LAST_NM,' +
                     'DOB,GNDR,RSDNC_STATE,IMUN_RGSTRY_STATE,VCCN_GRP,VCCN_ADMN_DT,DOSE_SERIES,' +
                     'BRND_NM,DOSE_SIZE,RCTN\n')

    fileOutputNotFound.write('MEAS_YR,MEMB_LIFE_ID_SKEY,MEMB_LIFE_ID,MEMB_FRST_NM,MEMB_LAST_NM,MEMB_SUFFIX,' +
                             'DOB,GNDR,RSDNC_STATE,IMUN_RGSTRY_STATE,VCCN_GRP,VCCN_ADMN_DT,DOSE_SERIES,' +
                             'BRND_NM,DOSE_SIZE,RCTN\n')

    df = pd.read_excel(file)

    # create array of People objects and member ID
    peopleArray = []
    memberIdArray = []
    df.dropna()
    total = len(df)
    not_found = 0
    found = 0

    # assign each record in the data frame into Person class
    for i in range(total):
        measYr = str(df.loc[i, "MEAS_YR"])
        memberId = str(df.loc[i, "MEMB_LIFE_ID"])
        memberIdSkey = str(df.loc[i, "MEMB_LIFE_ID_SKEY"])
        fname = str(df.loc[i, "MEMB_FRST_NM"])
        lname = str(df.loc[i, "MEMB_LAST_NM"])
        lnameSuffix = str(df.loc[i, "MEMB_SUFFIX"])
        inputDate = str(df.loc[i, "DOB"])
        # If date is null then assign an impossible date
        if not inputDate:
            dob = '01/01/1900'
        if '-' in inputDate:
            dob = datetime.datetime.strptime(
                inputDate, "%Y-%m-%d %H:%M:%S").strftime('%m/%d/%Y')
        else:
            dob = datetime.datetime.strptime(
                str(df.loc[i, "DOB"]), '%m/%d/%Y').strftime('%m/%d/%Y')
        gender = str(df.loc[i, "GNDR"])

        stateRes = str(df.loc[i, "RSDNC_STATE"])
        meas = ''

        p = Person(measYr, memberId, memberIdSkey, fname, lname,
                   lnameSuffix, dob, gender, stateRes, meas)

        # append array
        m = df.loc[i, "MEMB_LIFE_ID"]

        if (m not in memberIdArray):
            peopleArray.append(p)

        memberIdArray.append(m)

    PATH = os.getcwd()+'/'+'chromedriver'
    driver = webdriver.Chrome(PATH)
    driver.get("https://viis.vdh.virginia.gov/VIIS/logon.do")

    # work on org code
    username = driver.find_element_by_name("orgCode")
    username.clear()
    username.send_keys(org)

    # work on login ID
    username = driver.find_element_by_name("username")
    username.clear()
    username.send_keys(user)

    # work on password
    password = driver.find_element_by_name("password")
    password.clear()
    password.send_keys(pw)

    # work on getting to home page - where loop will start - log on button
    driver.find_element_by_xpath(
        "/html/body/table/tbody/tr[1]/td[1]/form/table[1]/tbody/tr[5]/td/input").click()

    # work on view client report button, wait 0.5 for result to show up
    time.sleep(0.25)
    driver.find_element_by_class_name('xMenuArea').click()

    for n in range(total):
        p = peopleArray[n]
        recordToWrite = ''
        print('Looking up: ' + str(n)+' ' +
              p.getLastName() + ', ' + p.getFirstName())
        MeasYr = p.getMeasYr()
        MemberIdSkey = p.getMemberIdSkey()
        MemberId = p.getMemberId()
        Fname = p.getFirstName()
        Lname = p.getLastName()
        DOB = str(p.getDateOfBirth())
        Gender = p.getGender()
        StateRes = p.getStateRes()

        try:
            children = immunte(Fname, Lname, DOB, Gender, driver)
        except UnexpectedAlertPresentException:
            time.sleep(0.5)
            children = immunte(Fname, Lname, DOB, Gender, driver)

        if children == []:
            not_found += 1
            recordToWrite = MeasYr+','+MemberIdSkey+','+MemberId+',' + Fname + \
                ','+Lname + ',' + ' ' + ','+DOB+','+Gender+','+StateRes+','+'VA'
            fileOutputNotFound.write(recordToWrite + '\n')
        elif children != []:
            found += 1
            for x in range(len(children)):
                data_element = children[x].split(",")

                if is_date(data_element[1]) and is_date(data_element[3]):
                    children[x] = ''
                elif is_date(data_element[1]) and data_element[2] == 'NOT' and data_element[3] == 'VALID':
                    children[x] = ''
                elif is_date(data_element[1]) and is_date(data_element[3]) == False:
                    if data_element[5] != 'No':
                        data_element[4] = data_element[5]
                        data_element[5] = ''
                        children[x] = ','.join(data_element[0:6])
                    else:
                        data_element[5] = ''
                        children[x] = ','.join(data_element[0:6])
                else:
                    children[x] = ''

            for x in range(len(children)):
                if children[x] != '':
                    recordToWrite = MeasYr+','+MemberIdSkey+','+MemberId+',' + \
                        Fname+','+Lname + ','+DOB+','+Gender+','+StateRes+','+'MD'
                    recordToWrite = recordToWrite+','+children[x]
                    fileOutput.write(recordToWrite + '\n')
        n = +1

    fileOutput.close()
    fileOutputNotFound.close()

    print('\n--------------------------------OUTPUT--------------------------------')
    print("Script completed.")
    print("There are "+str(total)+" members in the original lookup list provided.")
    print("There are "+str(found) +
          " members were found with records on the VA immunization website.")
    print("There are "+str(not_found) +
          " members were not found on the VA immunization website.\n")
    print('Files saved: \n' + fileOutputName + '\n' + fileOutputNameNotFound)
    print('\n----------------------------------------------------------------------\n')
##############################################################################


main()
