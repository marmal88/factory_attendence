# NON - Mac compatible version 001
"""
This script allows user to choose different items from a menu related to tracking employee's attendance coupled with the duration of work to compute overtime pay as well as to monitor absentee report. 

Author: NTU/NBS/MSBA 2021/Group A8 

This script requires the following dependencies datetime, os, csv, numpy and pandas to be installed.
This script should be run from "Group_Project" folder.

This script contains the following functions: 
     *menu - available program options
     *main - main function of the script 
     *inputdate - prompt user for date entry 
     *inputtime - prompt user for time entry 
     *validdate - validate inputdate
     *validtime - validate inputtime
     *inputqrcode - prompt user for Token ID
     *ScanToken - scanning and validation for each QRCode and Datetime
     *SetTokenProfile - add in new employee information and / or update existing one
     *findfiles - locate files in INOUT folder 
     *MergeIOFiles - generate monthly scanning record
     *OverTimeReport - generate overtime report based on a prompted date
     *AbsentReport - generate absentee report based on a prompted date

This script assumes following:
    1. An existing "employees.csv" file stored in "Group_Project" folder. 
    2. QR code comprises of 4 digits only.
    3. Employee ID starts with S followed by 4 digits. 
    4. This script does not validate the completeness of employee email address.

"""
import datetime as dt
import pandas as pd
import numpy as np
import os
import csv

# Part 03 of Project
def menu():
    """
    Display available options for users to choose from.  
    """
    while True:
        menuitems = '''
                ***** Factory Attendance *****
                S: Scan Token
                T: Set Token Profile
                M: Merge Input/ Output Files
                O: Over Time Report
                A: Absent Report
                Q: Quit
                '''
        print(menuitems)
        choice = input("Please choose one of the above options ").replace(" ", "").upper()
        if choice in ['S', 'T', 'M', 'O', 'A', 'Q']:
            return choice
        else:
            print("This is not a valid option")

# As Part 03 of Project
def main():
    """
    Direct user to selected menu option. 
    """
    while True:
        choice = menu()
        if choice == 'S':
            ScanToken()
        elif choice == 'T':
            SetTokenProfile()
        elif choice == 'M':
            MergeIOFiles()
        elif choice == 'O':
            OverTimeReport()
        elif choice == 'A':
            AbsentReport()
        else:
            break

def inputdate():
    """
    Prompt user for date entry. 
    """
    while True:
        date = input("Enter date (YYYY-MM-DD): ").strip()
        if len(date) > 10:
            print("You have entered more digits than necessary.")
        elif len(date) != 10:
            print("You have entered insufficient digits.")
        else:
            qryear1 = date[0:4]
            qrmonth1 = date[5:7]
            qrday1 = date[8:]
            vdate = validdate(qryear1, qrmonth1, qrday1)
            present = dt.datetime.now()
            if vdate == False:
                print("Date entered is not valid.")
            elif dt.date(int(qryear1), int(qrmonth1), int(qrday1)) > present.date() and vdate == True:
                print("Date entered needs to be before today.")
            else:
                qryear = int(qryear1)
                qrmonth = int(qrmonth1)
                qrday = int(qrday1)
                return qryear, qrmonth, qrday

def inputtime():
    """
    Prompt user for time entry. 
    """
    while True:
        time = input("Enter simulated in/out time (HH:MM): ").strip()
        if time[2:3] != ':':
            print("Please include ':' in time format")
        elif len(time) > 5:
            print("You have entered more digits than necessary.")
        elif len(time) != 5:
            print("You have entered insufficient digits.")
        else:
            qrhour1 = time[0:2]
            qrmin1 = time[3:]
            vtime = validtime(qrhour1, qrmin1)
            if vtime == False:
                print("Time entered is not valid.")
            else:
                qrhour = int(qrhour1)
                qrmin = int(qrmin1)
                return qrhour, qrmin

def validdate(year, month, day):
    """
    Validate date entry. 
    """
    try:
        dt.date(int(year), int(month), int(day))
        vdate = True
    except ValueError:
        vdate = False
    return vdate

def validtime(hour, min):
    """
    Validate time entry. 
    """
    try:
        dt.time(int(hour), int(min))
        vtime = True
    except ValueError:
        vtime = False
    return vtime

def inputqrcode():
    """
    Prompt and validate Token ID.
    """
    while True:
        qrcode = input("Enter simulated Token ID data (4-digits): ").strip()
        if qrcode == 'Q':
            return qrcode
        elif qrcode.isdigit() == False:
            print("Please enter a numeric TokenID")
        elif len(qrcode) > 4 or len(qrcode) < 4:
            print("Please key in exactly 4-digit tokenID")
        else:
            return qrcode

# Part 04 of Project
def ScanToken():
    '''
    Performs scanning and validation for each QRCode and Datetime.
    Output of all the different scans are consolidated by day and saved into a IN/OUT file for that day
    In file for timings before 1pm, OT file for timings after 1pm
    '''
    # Part 07 of Project
    if os.path.basename(os.getcwd()) == 'Group_Project':  # Check if current folder is still Group_Project 
        if os.path.exists(".\INOUT") == True:# Check if INOUT folder exists
            pass
        else:
            os.makedirs(".\INOUT")  # If INOUT folder does not exist create INOUT

    while True:
        qrcodes = []
        qrtimes = []
        qrdates = []
        qrcode = inputqrcode()
        if qrcode != 'Q':
            qryear, qrmonth, qrday = inputdate()
            qrhour, qrmin = inputtime()
        else:
            break

        qrcodes.append(qrcode)
        qrdates.append(dt.date(qryear, qrmonth, qrday))
        qrtimes.append(dt.time(qrhour, qrmin))

        data = {'Date': qrdates, 'Time': qrtimes, 'Token ID': qrcodes}
        df = pd.DataFrame(data)
        df1 = df.copy()
        df1['Date'] = pd.to_datetime(df1['Date'])
        df1["Date"].dt.strftime("%y/%m/%d")
        df1["Year"] = df1['Date'].dt.year
        min_year = df1["Year"].max()
        max_year = df1["Year"].min()
        df1["Month"] = df1['Date'].dt.month
        df1["Day"] = df1['Date'].dt.day
        df1['Hour'] = df1['Time'].apply(lambda x: x.hour)

        # As Part 06 of Project
        for year in range(min_year, max_year + 1):
            for month in range(1, 13):
                for day in range(1, 32):
                    df2 = df1[(df1["Year"] == year) & (df1["Month"] == month) & (df1["Day"] == day) & (df1["Hour"] < 13)]
                    if len(df2) != 0: #write csv files only for dates with scanning entries                       
                        df2 = df2.drop(["Year", "Month", "Day", "Hour"], axis=1)
                        df2["Time"] = df2["Time"].map(lambda t: t.strftime('%H:%M'))
                        fname = f"INOUT\\IN_{str(year).zfill(4)}{str(month).zfill(2)}{str(day).zfill(2)}.csv"
                        # append new entries in case the file for a day with existing record
                        if os.path.exists(fname):
                            df2.to_csv(fname, mode='a', header=False, encoding="utf-8", index=False)
                        else:
                            df2.to_csv(fname, encoding='utf-8', index=False)

                    df3 = df1[(df1["Year"] == year) & (df1["Month"] == month) & (df1["Day"] == day) & (df1["Hour"] >= 13)]
                    if len(df3) != 0:#write csv files only for dates with scanning entries
                        df3 = df3.drop(["Year", "Month", "Day", "Hour"], axis=1)
                        df3["Time"] = df3["Time"].map(lambda t: t.strftime('%H:%M'))
                        fname = f"INOUT\OT_{str(year).zfill(4)}{str(month).zfill(2)}{str(day).zfill(2)}.csv"
                        # append new entries in case the file for a day with existing record
                        if os.path.exists(fname):
                            df3.to_csv(fname, mode='a', header=False, encoding="utf-8", index=False)
                        else:
                            df3.to_csv(fname, encoding='utf-8', index=False)


# As Part 05 of Project
def newprofile(ID, csvName, newName1, newMobileNumber, newEMail, tokenID):
    """
    set up new employee file with information given based on function parameters. 
    """
    csvList = []
    with open(csvName, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            csvList.append(row)
        dfCsv = pd.DataFrame(csvList[1:], columns=csvList[0])
        newName = newName1

    while True:
        if newMobileNumber[0] == '8' or '9':
            if len(newMobileNumber) == 8:
                if newMobileNumber.isnumeric():
                    dfCsv.loc[dfCsv['EmployeeID'] == ID, 'MobileNumber'] = newMobileNumber
                    break

    while True:  # enter new mobile EMail
        if newEMail.count('@') == 1:
            dfCsv.loc[dfCsv['EmployeeID'] == ID, 'EMail'] = newEMail
            break

    while True:  # enter new tokenID
        if str(tokenID).isnumeric():  # prevent from returning "Q"
            break

    addLine = ID + "," + newName + "," + newMobileNumber + "," + newEMail + "," + str(tokenID)
    with open(csvName, "a", encoding="utf-8", newline='') as newEmployee:
        writer = csv.writer(newEmployee)
        writer.writerow(addLine.split(","))

    return print("Successfully added into employees.csv")


# As Part 05 of Project
def SetTokenProfile():
    """
    Performs reading and adding employee information. Prompts user for Employee ID.
    For new employee, new information is added. For existing employee, existing information is displayed.
    """
    if os.path.basename(os.getcwd()) != 'Group_Project':  # Check if current folder is still Group_Project
        os.chdir('..')  # Assume here that folder is in INOUT
    else:
        pass

    if 'employee.csv' in os.listdir():
        csvName = os.getcwd() + '\employees.csv'
    else:
        columns = ['EmployeeID', 'Name', 'MobileNumber', 'EMail', 'TokenID']
        employees = pd.DataFrame(columns=columns)
        employees.to_csv('employees.csv', index_label=False)
        csvName = os.getcwd() + '\employees.csv'

    while True:  # enter a valid ID
        ID = input('Please enter your employee ID(S+4-digit):').capitalize()
        if ID[0] != 'S':
            print("Please enter ID begins with 'S'")
        if not ID[1:].isnumeric() or len(ID[1:]) != 4:
            print("Please enter 4-digit after 'S'")
        else:
            break

    dfe = pd.read_csv("employees.csv")
    ids = dfe[dfe['EmployeeID'] == ID]
    if len(ids) > 0:
        print(ids.to_string(index=False))
        updtpro = input("Do you want to update this profile? [Y/N] ").replace(" ", "")
        if updtpro.upper() == 'Y':
            newName = input('Please enter your name:').capitalize()
            newMobileNumber = input("Please enter your mobile number(8 digits begin with '8' or '9'):")
            newEMail = input('Please enter your EMail(must have @):')
            tokenID = inputqrcode()
            newprofile(ID, csvName, newName, newMobileNumber, newEMail, tokenID)
    else:
        print("Employee Number does not exist please enter new data.")
        newName = input('Please enter your name:').capitalize()
        newMobileNumber = input("Please enter your mobile number(8 digits begin with '8' or '9'):")
        newEMail = input('Please enter your EMail(must have @):')
        tokenID = inputqrcode()
        newprofile(ID, csvName, newName, newMobileNumber, newEMail, tokenID)

# As Part 08 of Project
def findfiles(filenamein):
    '''
    Used to filter for files in INOUT
    '''
    files = []
    for file in os.listdir(os.getcwd()):
        files.append(file)

    found = []
    for f in files:
        if (f"IN_{filenamein}" in f) or (f"OT_{filenamein}" in f):
            found.append(1)
        else:
            found.append(0)

    if any(found) == True:
        return True
    else:
        return False

# As Part 08 of Project
def MergeIOFiles():
    '''
    Takes user input of Month/Year and returns a csv file with employee's working time for the month.

    All files under the specified directory are read and sorted based on file name pre-fix "In" or "Out".
    Duplicates in daily record are removed to ensure each employee has only one record of scan in/out on a particular day.
    Combined and cleaned monthly IN/OUT records (mergein/mergeout) are then merged to form a monthly record of employee's scan-in/out status (merget).

    Records with scan-in time but no scan-out time are assigned scan-out time at 14:00 on the day.
    Records with scan-out time but no scan-in time are assigned scan-in time at 08:00 on the day.
    :return: Returns a csv file showing employee's working time on each day for the requested month.
    :rtype: csv file
    '''
    # Ensure directory is in INOUT Folder
    if 'INOUT' in os.listdir():
        os.chdir(os.path.join(os.getcwd(), 'INOUT'))
    else:
        pass

    # Check for month year provided
    while True:
        yearmonth = input("Please enter the year-month (YYYY-MM):").replace(" ", "")
        year = yearmonth[:4]
        month = yearmonth[5:]
        vdate = validdate(year, month, str(1))
        filenamein = f"{str(year)}{str(month).zfill(2)}"
        ffile = findfiles(filenamein)
        if vdate==True and ffile==True:
            break
        else:
            print("Date is either not valid or file does not exist")

    datain = []
    dataot = []
    for file in os.listdir(os.getcwd()):
        if f"IN_{filenamein}" in file:
            fin = pd.read_csv(file)
            # remove duplicates for multiple scan-in on the day
            fin.sort_values(["Time"], inplace=True)
            # to keep first record on the same day
            fin.drop_duplicates(subset='Token ID', keep='first', inplace=True)
            datain.append(fin)
        if f"OT_{filenamein}" in file:
            fout = pd.read_csv(file)
            # remove duplicates for multiple scan-out on the day
            fout.sort_values(["Time"], inplace=True)
            # to keep last record on the same day
            fout.drop_duplicates(subset='Token ID', keep='last', inplace=True)
            dataot.append(fout)

    mergein = pd.concat(datain, ignore_index=True)
    mergeout = pd.concat(dataot, ignore_index=True)
    mergein["Token ID"] = mergein["Token ID"].astype("object")
    merget = pd.merge(mergein, mergeout, on=['Date', 'Token ID'], how='outer', suffixes=('_IN', '_OUT'))

    # for available scan-in but no scan-out cases and vice versa
    merget['Time_IN'].fillna('12:59', inplace=True)  # Assume earliest time to come into factory
    merget['Time_OUT'].fillna('14:00', inplace=True)  # Assume earliest time to leave factory

    # formatting datetime
    merget['Time_INf'] = merget['Date'] + ' ' + merget['Time_IN']
    merget['Time_INf'] = pd.to_datetime(merget['Time_IN'])
    merget['Time_OUTf'] = merget['Date'] + ' ' + merget['Time_OUT']
    merget['Time_OUTf'] = pd.to_datetime(merget['Time_OUT'])

    # filter hours and mins
    merget["Hour_In"] = merget['Time_INf'].dt.hour
    merget["Min_In"] = merget['Time_INf'].dt.minute
    merget["Hour_Out"] = merget['Time_OUTf'].dt.hour
    merget["Min_Out"] = merget['Time_OUTf'].dt.minute

    # calculate total working hours on the day
    merget["Hrs"] = np.where(merget["Min_Out"] >= merget['Min_In'], merget["Hour_Out"] - merget["Hour_In"],
                             (merget["Hour_Out"] - merget["Hour_In"] - 1))
    merget['Mins'] = np.where(merget["Min_Out"] >= merget['Min_In'], merget["Min_Out"] - merget["Min_In"],
                              (60 - (merget["Min_In"] - merget["Min_Out"])))

    merget = merget.drop(["Min_Out", "Min_In", "Hour_Out", "Hour_In", "Time_INf", "Time_OUTf"], axis=1)

    # update table formatting
    merget.columns = ['Date', 'In Time', 'Token ID', 'Out Time', "Hrs", "Mins"]
    m = merget[['Date', 'In Time', 'Out Time', 'Token ID', "Hrs", "Mins"]]
    m = m.sort_values(['Date', "In Time", "Out Time"])

    filenameout = f"MG_{str(year)}{str(month).zfill(2)}.csv"

    os.chdir('..')  # File ends in Group_Project Folder
    m.to_csv(filenameout, encoding='utf-8', index=False)  # added index = false

# As Part 09 of Project
def OverTimeReport():
    '''
    This is an overtime report, displays employee overtime hours by day and saves into CSV
    '''
    # Ensure correct file folder
    if os.path.basename(os.getcwd()) != 'Group_Project':  # Check if current folder is still Group_Project
        os.chdir('..') # Assume here that folder is in INOUT
    else:
        pass
    
    while True: 
        
        year, month, day = inputdate()
        pdate = dt.date(year, month, day)
    
        # finding the corresponding MG_YYYYMM.csv
        fileopen = str("MG_") + str(year) + str(month).zfill(2)
        if not ".csv" in fileopen:
            fileopen += ".csv"
        
        files = []
        for file in os.listdir(os.getcwd()):
            files.append(file)

        found = []
        for f in files:
            if fileopen in f:
                found.append(1)
            else:
                found.append(0)
    
        if any(found) == True:
            break
        else:
            print("Date is either not valid or file does not exist")

    dfmg = pd.read_csv(fileopen)
    dfmg = dfmg.loc[dfmg['Date'] == str(pdate)]

    # reading the employees.csv file to merge both files together
    dfe = pd.read_csv("employees.csv")
    dfe['TokenID'] = dfe['TokenID'].astype('object')
    dfe.drop_duplicates(keep='last', subset='EmployeeID', inplace=True)
    dfm = pd.merge(dfmg, dfe, left_on='Token ID', right_on='TokenID')

    # filtering out the date as per userinput
    dfmdate = dfm[dfm['Date'] == str(pdate)]

    # finding the total duration worked in minutes
    dfmdate['Duration'] = (dfmdate['Hrs'] * 60) + dfmdate['Mins']

    # finding the overtime duration
    dfmdate = dfmdate[dfmdate['Duration'] >= ((9 * 60) + 15)]  # we add 15 because duration only qualifies if he stays at least 9 hours and 15 mins at work.
    dfmdate['Overtime in mins'] = dfmdate['Duration'] - (9 * 60)  # computation overtime is the excess time after 9 hours

    # creation of 'Work' column
    dfmdate['Work'] = dfmdate['Hrs'].astype(str) + " Hours " + dfmdate['Mins'].astype(str) + " Mins "

    # generation of overtime report
    print("Over Time List For {}".format(pdate))
    dfreport = dfmdate[['EmployeeID', 'Name', 'Work', 'Overtime in mins']]
    filename = f"Daily_Overtime_report_{pdate}.csv"
    pd.set_option('display.max_rows', None)
    if len(dfreport) > 0:
        print(dfreport.reset_index(drop=True))
    else:
        print('There are no employees clocking overtime today')
    dfreport.to_csv(filename, encoding='utf-8', index=False)

# As Part 10 of Project
def AbsentReport():
    """
    This is an absentee report, displays employees who are absent by day and saves into CSV
    """
    # Ensure correct file folder
    if os.path.basename(os.getcwd()) != 'Group_Project':  # Check if current folder is still Group_Project
        os.chdir('..') # Assume here that folder is in INOUT
    else:
        pass
    
    while True: 
            
            year, month, day = inputdate()
            user_date = dt.date(year, month, day)
        
            # finding the corresponding MG_YYYYMM.csv
            csv_file = "MG_" + str(year) + str(month).zfill(2) + ".csv"
            
            files = []
            for file in os.listdir(os.getcwd()):
                files.append(file)
    
            found = []
            for f in files:
                if csv_file in f:
                    found.append(1)
                else:
                    found.append(0)
        
            if any(found) == True:
                break
            else:
                print("Date is either not valid of file does not exist")

    # get dataframe and merge both files based on TokenID
    df_inout = pd.read_csv(csv_file)
    df_ee = pd.read_csv("employees.csv")
    df_ee['TokenID'] = df_ee['TokenID'].astype('object')
    df_ee.drop_duplicates(keep='last', subset='EmployeeID', inplace=True)
    dfM1 = pd.merge(df_ee, df_inout, left_on="TokenID", right_on="Token ID", how="outer")
    # this will contain all employees with corrs TokenID scanned and employees w no Token recorded at work

    # filter list of employees working on user input date
    df2 = dfM1[dfM1["Date"] == str(user_date)]

    # extract list of employees who are not working on user input date
    data = df_ee[~df_ee.EmployeeID.isin(df2.EmployeeID)]

    # generation of absent report
    print("Absent List For {}".format(user_date))
    a_report = data[["EmployeeID", "Name"]].dropna()
    filename = f"Daily_Absent_report_{user_date}.csv"
    pd.set_option('display.max_rows', None)
    if len(a_report) > 0:
        print(a_report.reset_index(drop=True))
    else:
        print("There are no absentees today")
    a_report.to_csv(filename, encoding='utf-8', index=False)

if __name__=='__main__':
    main()