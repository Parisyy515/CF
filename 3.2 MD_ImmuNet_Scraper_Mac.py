
# Filename: MD_ImmuNet_Scraper_Mac.py
# Author: Zheng Guo
# Date: 10-16-2020
# Purpose: Scraping member's immunization registration information from the MD Immunet site based on a given list of members.
# Class list: - Person (measYr, memberId, memberIdSkey, fname, lname, lnameSuffix, dob, gender, stateRes, meas)

# Functions:
# - is_date(string)
# - immunte(Fname, Lname, DOB, Gender, user, pw)

# User Input:
# - First input : MLQM_Immun_Regs_Lkup_SAMPLE.xlsx
#             xlsx file contains the following information in order below:
#             meas_yr,memb_life_id,memb_life_id_skey,memb_frst_nm,memb_last_nm,memb_nm_suffix,memb_dob,gender,state,meas
# - Second input : User name
# - Third input: Password

# Output:
# - First output: HEDIS_MD_Immun_Records_Found_YYYY_MM_DD.csv
# - Second output: HEDIS_MD_Immun_Records_Not_Found_YYYY_MM_DD.csv
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

    # work on patient search button
    driver.find_element_by_xpath("//*[@id='editVFCProfileButton']").click()

    # work on last name
    lastname = driver.find_element_by_id("txtLastName")
    lastname.clear()
    lastname.send_keys(Lname)

    # work on first name
    firstname = driver.find_element_by_id("txtFirstName")
    firstname.clear()
    firstname.send_keys(Fname)

    # work on birth date
    birthdate = driver.find_element_by_id("txtBirthDate")
    birthdate.clear()
    birthdate.send_keys(DOB)

    # work on advanced search button to input gender
    driver.find_element_by_xpath(
        "//*[@id='queryResultsForm']/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td[5]/input").click()

    # work on gender selection button
    obj = Select(driver.find_element_by_name("optSexCode"))
    if Gender == 'M':
        obj.select_by_index(2)
    elif Gender == 'F':
        obj.select_by_index(1)
    else:
        obj.select_by_index(3)

    # work on search button
    driver.find_element_by_name("cmdFindClient").click()

    # two scenarios could emerge as a search result: 1, no patient found 2, the patient found
    if "No patients were found for the requested search criteria" in driver.find_element_by_id("queryResultsForm").text:
        al = []

    elif "Patient Demographics Patient Immunization History" in driver.find_element_by_id("queryResultsForm").text:

        # work on patient immunization button
        driver.find_element_by_xpath(
            "//*[@id='queryResultsForm']/table[2]/tbody/tr[2]/td[2]/span/label").click()

        # work on patient last name button
        driver.find_element_by_id("redirect1").click()

        # work on getting rid of people who opt out of the site - header
        header = driver.find_elements_by_class_name("large")[1].text

        if "Access Restricted" in header:
            print(Fname+' '+Lname+' '+" Opt out")
            al = []

        elif "Patient Information" in header:
            # find the first line
            first = driver.find_element_by_xpath(
                "//*[@id='container']/table[3]/tbody/tr/td[2]/table[2]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr[5]/td[1]").text
            if (first == None):
                al = []

            else:
                even = driver.find_elements_by_class_name("evenRow")
                odd = driver.find_elements_by_class_name("oddRow")
                o = []
                e = []

                for value in odd:
                    o.append(value.text)
                for value in even:
                    e.append(value.text)

                length = len(o)
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

    # work on returning to home page
    driver.find_element_by_xpath(
        "//*[@id='headerMenu']/table/tbody/tr/td[2]/div/a").click()

    return al


def main():
    # Welcome message and input info
    print('\nThis is the web scraper for the MaryLand Immunization Record Website.')
    print('You will be prompted to type in a file name and username/password.')
    print('If you need to exit the script and stop its process press \'CTRL\' + \'C\'.')
    file = input("\nEnter file name: ")
    user = input("\nEnter MDImmnet username: ")
    pw = input("\nEnter MDImmnet password: ")

    date = str(datetime.date.today())

    # output file
    fileOutputName = 'HEDIS_MD_Immun_Records_Found_' + \
        date.replace('-', '_') + '.csv'
    fileOutputNameNotFound = 'HEDIS_MD_Immun_Records_Not_Found_' + \
        date.replace('-', '_') + '.csv'

    fileOutput = open(fileOutputName, 'w')
    fileOutputNotFound = open(fileOutputNameNotFound, 'w')

    fileOutput.write('MEAS_YR,MEMB_LIFE_ID_SKEY,MEMB_LIFE_ID,MEMB_FRST_NM,MEMB_LAST_NM,' +
                     'DOB,GNDR,RSDNC_STATE,IMUN_RGSTRY_STATE,VCCN_GRP,VCCN_ADMN_DT,DOSE_SERIES,' +
                     'BRND_NM,DOSE_SIZE,RCTN\n')

    fileOutputNotFound.write('MEAS_YR,MEMB_LIFE_ID_SKEY,MEMB_LIFE_ID,MEMB_FRST_NM,MEMB_LAST_NM,MEMB_SUFFIX,' +
                             'DOB,GNDR,RSDNC_STATE,IMUN_RGSTRY_STATE,VCCN_GRP,VCCN_ADMN_DT,DOSE_SERIES,' +
                             'BRND_NM,DOSE_SIZE,RCTN\n')

    # If the file exists
    try:
        os.path.isfile(file)
    except:
        print('File Not Found\n')

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
        measYr = str(df.loc[i, "#MEAS_YR"])
        memberId = str(df.loc[i, "MEMB_LIFE_ID"])
        memberIdSkey = str(df.loc[i, "MEMB_LIFE_ID_SKEY"])
        fname = str(df.loc[i, "MEMB_FRST_NM"])
        lname = str(df.loc[i, "MEMB_LAST_NM"])
        lnameSuffix = str(df.loc[i, "MEMB_NM_SUFFIX"])
        inputDate = str(df.loc[i, "MEMB_DOB"])
        # If date is null then assign an impossible date
        if not inputDate:
            dob = '01/01/1900'
        if '-' in inputDate:
            dob = datetime.datetime.strptime(
                inputDate, "%Y-%m-%d %H:%M:%S").strftime('%m/%d/%Y')
        else:
            dob = datetime.datetime.strptime(
                str(df.loc[i, "MEMB_DOB"]), '%m/%d/%Y').strftime('%m/%d/%Y')
        gender = str(df.loc[i, "GENDER"])
        stateRes = str(df.loc[i, "STATE_RES"])
        meas = str(df.loc[i, "MEAS"])

        p = Person(measYr, memberId, memberIdSkey, fname, lname,
                   lnameSuffix, dob, gender, stateRes, meas)

        # append array
        m = df.loc[i, "MEMB_LIFE_ID"]

        if (m not in memberIdArray):
            peopleArray.append(p)

        memberIdArray.append(m)

    # work on setting up driver for md immunet - mac forward slash/windows double backward slash
    PATH = os.getcwd()+'/'+'chromedriver'
    driver = webdriver.Chrome(PATH)
    driver.get("https://www.mdimmunet.org/prd-IR/portalInfoManager.do")

    # work on login ID
    username = driver.find_element_by_id("userField")
    username.clear()
    username.send_keys(user)

    # work on password
    password = driver.find_element_by_name("password")
    password.clear()
    password.send_keys(pw)

    # work on getting to home page - where loop will start
    driver.find_element_by_xpath(
        "//*[@id='loginButtonForm']/div/div/table/tbody/tr[3]/td[1]/input").click()

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
        children = immunte(Fname, Lname, DOB, Gender, driver)

        if children == []:
            not_found += 1
            recordToWrite = MeasYr+','+MemberIdSkey+','+MemberId+',' + Fname + \
                ','+Lname + ',' + ' ' + ','+DOB+','+Gender+','+StateRes+','+'MD'
            fileOutputNotFound.write(recordToWrite + '\n')
        elif children != []:
            found += 1
            for x in range(len(children)):
                data_element = children[x].split(",")

                # if the admin date is not valid, or the brand is not valid skip the records, clean data on the dosage and reaction field
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
          " members were found with records on the MD immunization website.")
    print("There are "+str(not_found) +
          " members were not found on the MD immunization website.\n")
    print('Files saved: \n' + fileOutputName + '\n' + fileOutputNameNotFound)
    print('\n----------------------------------------------------------------------\n')
##############################################################################


main()
