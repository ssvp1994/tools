'''This program need only Excel files with .xls format, any other format won't work and specifically desined for weekly reports
of E.Tools - T2 (KARSUN SOLUTIONS)'''

import os
import datetime
import xlwt
import xlrd
import mysql.connector
from pandas import ExcelFile
import pandas as pd


def top_row(final):  # naming the indexes of the file
    '''Rename the top row of the sheet with required column names'''
    xls_file1 = ExcelFile(final, index=True)
    df = xls_file1.parse('Page 1')
    print("\nRenaming")
    # naming the indexes of the file
    df4 = df.rename(
        columns={"Unnamed: 0": 'Number', "Unnamed: 1": 'Priority', "Unnamed: 2": 'Opened', "Unnamed: 3": 'Definition',
                 "Unnamed: 4": 'Value', "Unnamed: 5": 'Expert_Assigned', "Unnamed: 6": 'Created',
                 "Unnamed: 7": 'End_time', "Unnamed: 8": 'Resolved', "Unnamed: 9": 'New_Resolved',
                 "Unnamed: 10": 'Closed', "Unnamed: 11": 'Main_file'})
    df4.to_excel(final, sheet_name='Page 1', index=False)
    return


def combing_files(path,newfinal):
    '''Combines all rows in all xls files of a given folder'''
    outrow = 2 #used as the parametre that counts the rows in final file
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)

    for x in os.listdir(path):
        os.rename(path + '/' + x, path + '/' + x.replace(" ", ""))
    y = sorted(os.listdir(path))
    for file in y:
        print("\nprocessing :"+file)
        open = xlrd.open_workbook(path + r'/' + file)
        insheet=open.sheet_by_name("Page 1")

        for row in range(1, ((insheet.nrows))):
            if len(str(insheet.cell_value(row,0)))>0:
                for col in range(0, 5):
                    outsheet.write(outrow, col, insheet.cell_value(row, col))#writes to specific col and row in the out file
                outsheet.write(outrow, 6, insheet.cell_value(row, 5))
                outsheet.write(outrow, 8, insheet.cell_value(row, 6))
                outsheet.write(outrow, 10, insheet.cell_value(row, 7))
                outsheet.write(outrow, 11, file)
                outrow += 1
    wb.save(newfinal)
    return

def drop_duplicates(final, keep):
    '''Drop's the Duplicate rows as some files have same repeated rows'''
    xls_file = ExcelFile(final, index=False)
    df = xls_file.parse('Page 1')
    print("\nDropping duplicates")
    df4 = df.drop_duplicates(subset=['Number', 'Expert_Assigned', 'Opened', 'Definition', 'Value', 'Created'],
                             keep=keep)
    df5 = df4.sort_values(['Number', 'Created', 'Definition', 'Expert_Assigned'], ascending=[True, True, False, True])
    df5.to_excel(final, sheet_name='Page 1', index=False)
    return


def end_time(file):
    '''print end time for every row, simple calculation - end time of row 1 = begin time of row 2'''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")  # loading the file to insheet

    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))

    for row in range(1, ((insheet.nrows) - 1)):
        for col in range(0, 1):
            calc = insheet.cell_value(row, 5)
            if insheet.cell_value(row, 3) == "Assigned to Duration" and len(str(calc)) == 0:
                outsheet.write(row, 5, insheet.cell_value(row, 4))
                outsheet.write(row, 4, insheet.cell_value(row - 1, 4))
            if insheet.cell_value(row, 0) == insheet.cell_value(row + 1, 0):
                outsheet.write(row, 7, insheet.cell_value(row + 1, 6))
                outsheet.write(row + 1, 7, 72733)

    wb.save(file)
    return


def new_resolved_cols(file):  # printing new values where the facts were wrong by few seconds in the resolved column
    '''Some tickets in the files have a resolved state which was started prior to assignment of the ticket to some one - these kind of tickets were wrong mostly by 10 secs'''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")
    print("\ncorrecting time stamps")
    for row in range(0, insheet.nrows):

        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
        if insheet.cell_value(row, 7) is "":
            outsheet.write(row, 7, 72733)

    for row in range(1, insheet.nrows):
        temp = (insheet.cell_value(row, 8))
        if len(str(temp)) > 0:
            outsheet.write(row, 9, temp)
    outsheet.write(0, 9, "New_Resolved")

    for row in range(1, ((insheet.nrows) - 1)):
        t1 = insheet.cell_value(row, 6)

        t3 = insheet.cell_value(row, 8)
        if len(str(t3)) > 0:
            t3f = float(t3)
            t1f = float(t1)

            t5 = t3f + 0.00014
            if ((insheet.cell_value(row, 0) == insheet.cell_value(row + 1, 0) or insheet.cell_value(row,
                                                                                                    0) == insheet.cell_value(
                    row - 1, 0)) and t5 > t1f and abs(t1f - t3f) < 0.00017):
                outsheet.write(row, 9, insheet.cell_value(row, 6))

                for sec in range(-10, 4):
                    if ((row + sec) <= insheet.nrows - 1):
                        if (insheet.cell_value(row, 0) == insheet.cell_value((row + sec), 0)):
                            outsheet.write((row + sec), 9, insheet.cell_value(row, 6))

        if row == ((insheet.nrows) - 1):
            if insheet.cell_value(row, 0) == insheet.cell_value(row - 1, 0):
                if len(str(t3)) > 0:
                    if t1f - t3f < 0.00017 and t1f - t3f > 0:
                        outsheet.write(row, 9, insheet.cell_value(row, 6))

    wb.save(file)
    return


def res_row(file):
    '''Reopened tickets'''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")
    print("\nPrinting Resolved rows")

    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))

    nrows = insheet.nrows
    for row in range(1, ((insheet.nrows) - 1)):
        t1 = insheet.cell_value(row, 5)
        t2 = insheet.cell_value(row, 6)
        t3 = insheet.cell_value(row, 9)
        if len(str(t3)) > 0:
            t3f = float(t3)
            t2f = float(t2)
            t5 = insheet.cell_value(row + 1, 6)
            t5f = float(t5)
            if len(str(t1)) > 0 and insheet.cell_value(row, 0) == insheet.cell_value(row + 1,
                                                                                     0) and t3f >= t2f and t3f < t5f:
                outsheet.write(nrows + 1, 3, "Resolved")
                outsheet.write(nrows + 1, 0, insheet.cell_value(row, 0))
                outsheet.write(nrows + 1, 1, insheet.cell_value(row, 1))
                outsheet.write(nrows + 1, 2, insheet.cell_value(row, 2))
                outsheet.write(nrows + 1, 4, insheet.cell_value(row, 4))
                outsheet.write(nrows + 1, 5, insheet.cell_value(row, 5))
                outsheet.write(nrows + 1, 6, insheet.cell_value(row, 9))
                outsheet.write(nrows + 1, 7, insheet.cell_value(row + 1, 6))
                outsheet.write(nrows + 1, 12, "Opened")
                outsheet.write(row, 7, insheet.cell_value(row, 9))
                nrows += 1

    wb.save(file)
    return


def correct_dept(newfinal):
    '''we changed the entire 'value' column with it been only department names and name of the person is shifted to Expert assigned  '''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(newfinal)
    insheet = open.sheet_by_name("Page 1")
    print("\nChecking department names")
    count = 0
    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    for all in range(1, ((insheet.nrows))):
        t1 = insheet.cell_value(all, 4)
        if (all + 3) < insheet.nrows:
            if len(t1) > 28:
                count += 1
                if len(insheet.cell_value(all - 1, 4)) < 29:
                    outsheet.write(all, 4, insheet.cell_value(all - 1, 4))
                elif len(insheet.cell_value(all - 1, 4)) > 29:
                    if len(insheet.cell_value(all - 2, 4)) < 29:
                        outsheet.write(all, 4, insheet.cell_value(all - 2, 4))
                    elif len(insheet.cell_value(all - 2, 4)) > 29:
                        if len(insheet.cell_value(all - 3, 4)) < 29:
                            outsheet.write(all, 4, insheet.cell_value(all - 3, 4))
                        elif len(insheet.cell_value(all - 3, 4)) > 29:
                            if len(insheet.cell_value(all - 4, 4)) < 29:
                                outsheet.write(all, 4, insheet.cell_value(all - 4, 4))
        if (all + 1) == insheet.nrows - 1:
            t1 = insheet.cell_value((all + 1), 4)
            if len(t1) > 28:
                if len(insheet.cell_value(all, 4)) < 29:
                    outsheet.write(all + 1, 4, insheet.cell_value(all, 4))
    print("\nFound " + str(count) + " errors")
    print("\nCorrecting " + str(count) + " errors")
    wb.save(newfinal)
    return


def end_state(newfinal):
    '''How the particular incident ended'''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(newfinal)
    insheet = open.sheet_by_name("Page 1")
    print("\nPrinting End States")
    nrows1 = insheet.nrows

    for row in range(0, insheet.nrows - 1):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    outsheet.write(0, 12, "End_State")
    for row in range(1, insheet.nrows - 1):

        if insheet.cell_value(row, 0) == insheet.cell_value(row + 1, 0):
            if insheet.cell_value(row, 1) != insheet.cell_value(row + 1, 1):
                outsheet.write(row, 12, "Priority changed")
            if insheet.cell_value(row, 3) == insheet.cell_value(row + 1, 3):
                outsheet.write(row, 12, insheet.cell_value(row + 1, 3))
            if insheet.cell_value(row, 3) != insheet.cell_value(row + 1, 3):
                outsheet.write(row, 12, insheet.cell_value(row + 1, 3))

        if insheet.cell_value(row, 0) != insheet.cell_value(row + 1, 0):
            t1 = insheet.cell_value(row, 9)
            t2 = insheet.cell_value(row, 6)
            if len(str(t1)) > 0:
                t1f = float(t1)
                t2f = float(t2)
                if t1f >= t2f:
                    outsheet.write(row, 12, "Resolved")
                    outsheet.write(nrows1 + 1, 3, "Resolved")
                    outsheet.write(nrows1 + 1, 0, insheet.cell_value(row, 0))
                    outsheet.write(nrows1 + 1, 1, insheet.cell_value(row, 1))
                    outsheet.write(nrows1 + 1, 2, insheet.cell_value(row, 2))
                    outsheet.write(nrows1 + 1, 4, insheet.cell_value(row, 4))
                    outsheet.write(nrows1 + 1, 5, insheet.cell_value(row, 5))
                    outsheet.write(nrows1 + 1, 6, insheet.cell_value(row, 9))
                    outsheet.write(nrows1 + 1, 7, 72733)
                    outsheet.write(nrows1 + 1, 12, "Resolved")
                    outsheet.write(row, 7, insheet.cell_value(row, 9))

                    nrows1 += 1

    wb.save(newfinal)
    return


def format_date(file):
    '''Excel sheet always stores the date in Excel format so when we directly look at the newly created report that doesnt make sense so we convert those format to readable format'''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")

    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    outsheet.write(0, 13, "Formatted_Created")
    outsheet.write(0, 14, "Formatted_End_time")
    outsheet.write(0, 15, "Formatted_Resolved")
    outsheet.write(0, 16, "Formatted_New_Resolved")
    outsheet.write(0, 17, "Formatted_Closed")

    for row in range(1, insheet.nrows):
        if len(str(insheet.cell_value(row, 6))) > 0:
            value13 = xlrd.xldate.xldate_as_datetime(float(insheet.cell_value(row, 6)), open.datemode)
            outsheet.write(row, 13, str(value13))
        if len(str(insheet.cell_value(row, 7))) > 0:
            value14 = xlrd.xldate.xldate_as_datetime(float(insheet.cell_value(row, 7)), open.datemode)
            outsheet.write(row, 14, str(value14))
        if len(str(insheet.cell_value(row, 8))) > 0:
            value15 = xlrd.xldate.xldate_as_datetime(float(insheet.cell_value(row, 8)), open.datemode)
            outsheet.write(row, 15, str(value15))
        if len(str(insheet.cell_value(row, 9))) > 0:
            value16 = xlrd.xldate.xldate_as_datetime(float(insheet.cell_value(row, 9)), open.datemode)
            outsheet.write(row, 16, str(value16))
        if len(str(insheet.cell_value(row, 10))) > 0:
            value17 = xlrd.xldate.xldate_as_datetime(float(insheet.cell_value(row, 10)), open.datemode)
            outsheet.write(row, 17, str(value17))

    wb.save(file)
    return


def to_excel_date(path):
    '''prints the date when the file was created'''

    def excel_date(date1):
        temp = datetime.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        delta = date1 - temp
        return (float(delta.days) + (float(delta.seconds) / 86400))

    listed = []
    for file in os.listdir(path):
        get = os.stat(path + r'/' + file).st_ctime
        convert = (datetime.datetime.fromtimestamp(get))
        listed.append(excel_date(convert))
    return listed


def time_diff(path, file):
    '''t'''
    maxed = max(to_excel_date(path))
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=False)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")
    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    outsheet.write(0, 18, "Time_Diff")
    for row in range(1, insheet.nrows):
        if len(str(insheet.cell_value(row, 7))) > 0:
            if insheet.cell_value(row, 7) == 72733:

                outsheet.write(row, 18, (maxed - insheet.cell_value(row, 6)))
            else:
                outsheet.write(row, 18, ((float(insheet.cell_value(row, 7))) - insheet.cell_value(row, 6)))
    wb.save(file)
    return


def aggregate(file):
    '''total time spent on the ticket useful in calculating the SLA'''
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=False)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")
    list1 = []
    dict1 = {}
    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    for row in range(1, ((insheet.nrows) - 1)):
        if insheet.cell_value(row, 0) == insheet.cell_value(row + 1, 0) and insheet.cell_value(row,
                                                                                               3) != "Resolved" and insheet.cell_value(
                row, 18) != '':
            list1.append(insheet.cell_value(row, 18))

            total = sum(list1)
        elif insheet.cell_value(row, 0) != insheet.cell_value(row + 1, 0) and insheet.cell_value(row,
                                                                                                 3) != "Resolved" and insheet.cell_value(
                row, 18) != '':
            list1.append(insheet.cell_value(row, 18))

            total = sum(list1)

        if insheet.cell_value(row, 0) != insheet.cell_value(row + 1, 0):
            list1 = []
            dict1[insheet.cell_value(row, 0)] = total

    if row == insheet.nrows - 2:

        row += 1
        if insheet.cell_value(row, 3) != "Resolved" and insheet.cell_value(row, 0) != None:
            total += insheet.cell_value(row, 18)
            dict1[insheet.cell_value(row, 0)] = total

    outsheet.write(0, 19, "Aggregate")
    for row in range(1, insheet.nrows):

        if insheet.cell_value(row, 3) != "Resolved":
            outsheet.write(row, 19, dict1[insheet.cell_value(row, 0)])
    wb.save(file)
    return


def holidays(startdate, enddate):
    ''' output's number of days with out fedaral holidays in the year 2016 2017 2018 when start and end date are passed through args'''
    days = range(int(startdate), int(enddate))
    weekends = []

    fedaral = {42005: "New Year's Day", 42023: "Birthday of Martin Luther King, Jr.", 42051: "Washington's Birthday",
               42149: "Memorial Day",
               42188: "Independence Day", 42254: "Labor Day", 42289: "Columbus Day", 42319: "Veterans Day",
               42334: "Thanksgiving Day", 42363: "Christmas Day"}
    fedaral.update(
        {42370: "New Year's Day", 42387: "Birthday of Martin Luther King, Jr.", 42415: "Washington's Birthday",
         42520: "Memorial Day",
         42555: "Independence Day", 42618: "Labor Day", 42653: "Columbus Day", 42685: "Veterans Day",
         42698: "Thanksgiving Day", 42730: "Christmas Day"})
    fedaral.update(
        {42737: "New Year's Day", 42751: "Birthday of Martin Luther King, Jr.", 42786: "Washington's Birthday",
         42884: "Memorial Day", 42920: "Independence Day", 42982: "Labor Day",
         43017: "Columbus Day", 43049: "Veterans Day", 43062: "Thanksgiving Day", 43094: "Christmas Day"})
    fedaral.update(
        {43101: "New Year's Day", 43115: "Birthday of Martin Luther King, Jr.", 43150: "Washington's Birthday",
         43248: "Memorial Day", 43285: "Independence Day", 43346: "Labor Day",
         43381: "Columbus Day", 43415: "Veterans Day", 43426: "Thanksgiving Day", 43459: "Christmas Day"})
    fedarallist = []
    if enddate > 43459:
        print("\nFedaral Holidays Calculation Expired, Please update it for precise results")

    for i in fedaral.keys():
        fedarallist.append(i)
    fedaralhol = [x for x in fedarallist if x >= float(startdate) and x <= float(enddate)]

    for day in days:
        if day % 7 == 0 or day % 7 == 1:
            weekends.append(day)
    # print("\nNumber of weekends in given range = %s" %len(weekends), "\nNumber of Fedaral in given range = %s" %len(fedaralhol))
    # print(abs(int(startdate)-int(enddate))-len(weekends)-len(fedaralhol))
    return abs(float(startdate) - float(enddate)) - len(weekends) - len(fedaralhol)


def timetaken(file):
    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=False)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")

    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    outsheet.write(0, 20, "Time_excluding_holidays")
    for row in range(1, insheet.nrows - 1):

        if insheet.cell_value(row, 0) == insheet.cell_value(row + 1, 0) and insheet.cell_value(row,
                                                                                               3) != "Resolved":
            outsheet.write(row, 20, (holidays((insheet.cell_value(row, 6)), (insheet.cell_value(row, 7)))))

    if row == insheet.nrows - 2:
        row += 1
        if insheet.cell_value(row, 3) != "Resolved":
            outsheet.write(row, 20, (
                holidays((insheet.cell_value(row, 6)), (insheet.cell_value(row, 7)))))

    wb.save(file)
    return


def testing(path, newfile):  # Testing block
    '''Test final data with the original data'''
    print("\n*********************************        TESTING      **************************************")
    open = xlrd.open_workbook(newfile)
    insheet1 = open.sheet_by_name("Page 1")
    set1 = set()
    for test in os.listdir(path):
        if len(str(test)) > 10:
            print(test)

            open2 = xlrd.open_workbook(path + '/' + str(test))
            insheet2 = open2.sheet_by_name("Page 1")
            for row in range(1, insheet2.nrows - 1):
                if insheet2.cell_value(row, 0) != insheet2.cell_value(row - 1, 0):
                    for search in range(1, insheet1.nrows - 1):
                        if test == insheet1.cell_value(search, 11):
                            if insheet1.cell_value(search, 0) != insheet1.cell_value(search - 1,
                                                                                     0) and insheet2.cell_value(row,
                                                                                                                0) == insheet1.cell_value(
                                search, 0) and (str(insheet2.cell_value(row, 0)) in set1) == False:
                                set1.add(insheet2.cell_value(row, 0))
                                if insheet2.cell_value(row, 2) != insheet1.cell_value(search, 2) or insheet2.cell_value(
                                        row, 5) != insheet1.cell_value(search, 6) or insheet2.cell_value(row,
                                                                                                         6) != insheet1.cell_value(
                                    search, 8):
                                    print("\ncheck failed for :" + str(
                                        insheet2.cell_value(row, 0)) + " in the file :" + str(test))
        print("\nFinished testing : " + str(test))

    return


def database_operations():
    '''the whole database operations are under this function'''
    fromdb = r"/var/www/html/cgi-bin/db_operations/fromdb.xls"  # path to folder to save from database
    database = mysql.connector.connect(host='gsar1.karsun-csb.com', port='59100', user='csbops',
                                       password='csbopspswd', db="csbops")
    cursor = database.cursor()

    def copy_from_db():

        table = "Fact_table"

        query = "SELECT * FROM %s;" % table

        cursor.execute(query)
        wb = xlwt.Workbook()
        outsheet = wb.add_sheet('Page', cell_overwrite_ok=True)
        header = cursor.column_names
        for col in range(0, len(header)):
            outsheet.write(0, col, header[col])

        for r, row in enumerate(cursor.fetchall()):
            for c, col in enumerate(row):
                outsheet.write(r + 1, c, str(col))

        wb.save(fromdb)

        open1 = xlrd.open_workbook(fromdb)
        insheet = open1.sheet_by_name("Page")
        outsheet1 = wb.add_sheet('Page 1', cell_overwrite_ok=True)
        for row in range(0, insheet.nrows):
            for col in range(0, insheet.ncols):
                outsheet1.write(row, col, insheet.cell_value(row, col))
        for row in range(1, insheet.nrows):
            for col in [2, 6, 7, 8, 9, 10]:
                if len(insheet.cell_value(row, col)) > 0:
                    outsheet1.write(row, col, float(insheet.cell_value(row, col)))
        wb.save(fromdb)

        cursor.close()
        database.commit()
        database.close()

    def remove_common_data():
        new_data = r"/var/www/html/cgi-bin/downloads/newcombined.xls"

        final = r"/var/www/html/cgi-bin/db_operations/To_db.xls"

        xls_file = ExcelFile(new_data, index=True)
        df = xls_file.parse('Page 1')
        xls_file2 = ExcelFile(fromdb, index=True)
        df2 = xls_file2.parse('Page 1')
        list1 = [df, df2, df2]
        concatenating = pd.concat(list1)

        new_rows = concatenating.drop_duplicates(
            subset=['Number', 'Expert_Assigned', 'Opened', 'Definition', 'Value', 'Created'],
            keep=False)
        new_rows.to_excel(final, sheet_name='Page 1', index=False)

    def copy_to_db():
        database = mysql.connector.connect(host='gsar1.karsun-csb.com', port='59100', user='csbops',
                                           password='csbopspswd', db="csbops")
        cursor = database.cursor()
        #file = r"/var/www/html/cgi-bin/db_operations/To_db.xls"
        file = r"/var/www/html/cgi-bin/downloads/newcombined.xls"

        sheet = xlrd.open_workbook(file).sheets()[0]

        print("\n\n Copying to database")
        cursor.execute("delete from Fact_table;")
        query = "INSERT INTO Fact_table (Number,	Priority,	Opened,	Definition,	Value,	Expert_Assigned,	Created,	End_time,	Resolved,	New_Resolved,	Closed,	Main_file,	End_State,Formatted_Created,Formatted_End_time," \
                "Formatted_Resolved,Formatted_New_Resolved,Formatted_Closed,Time_Diff,Aggregate,Time_excluding_holidays) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s,%s,%s,%s,%s,%s,%s,%s)"

        for r in range(1, sheet.nrows):
            Number = sheet.cell(r, 0).value
            Priority = sheet.cell(r, 1).value
            Opened = sheet.cell(r, 2).value
            Definition = sheet.cell(r, 3).value
            Value = sheet.cell(r, 4).value
            Expert_Assigned = sheet.cell(r, 5).value
            Created = sheet.cell(r, 6).value
            End_time = sheet.cell(r, 7).value
            Resolved = sheet.cell(r, 8).value
            New_Resolved = sheet.cell(r, 9).value
            Closed = sheet.cell(r, 10).value
            Main_file = sheet.cell(r, 11).value
            End_State = sheet.cell(r, 12).value
            Formatted_Created = sheet.cell(r, 13).value
            Formatted_End_time = sheet.cell(r, 14).value
            Formatted_Resolved = sheet.cell(r, 15).value
            Formatted_New_Resolved = sheet.cell(r, 16).value
            Formatted_Closed = sheet.cell(r, 17).value
            Time_Diff = sheet.cell(r, 18).value
            Aggregate = sheet.cell(r, 19).value
            Time_excluding_holidays = sheet.cell(r, 20).value

            values = (
                Number, Priority, Opened, Definition, Value, Expert_Assigned, Created, End_time, Resolved, New_Resolved,
                Closed,
                Main_file, End_State, Formatted_Created, Formatted_End_time, Formatted_Resolved, Formatted_New_Resolved,
                Formatted_Closed, Time_Diff, Aggregate, Time_excluding_holidays)

            cursor.execute(query, values)

            # Close the cursor
        cursor.close()

        # Commit the transaction
        database.commit()

        # Close the database connection
        database.close()

    #copy_from_db()
    #remove_common_data()
    copy_to_db()

    return


def dim_table(file):
    sheet = xlrd.open_workbook(file).sheets()[0]
    columns = []
    names = []

    for i in range(sheet.ncols):
        columns.append(sheet.cell_value(0, i) + " varchar(255)")
        names.append(sheet.cell_value(0, i))

    database = mysql.connector.connect(host='gsar1.karsun-csb.com', port='59100', user='csbops',
                                       password='csbopspswd', db="csbops")
    cursor = database.cursor()
    cursor.execute("show tables")
    for r, row in enumerate(cursor.fetchall()):
        for c, col in enumerate(row):
            if col=="dim_incident":
                cursor.execute("Drop table dim_incident;")

    cursor.execute("Create table dim_incident (%s);" % columns[0])

    for x in range(1, len(columns)):
        cursor.execute("ALTER TABLE dim_incident add(%s)" % columns[x])

    for record in range(1, sheet.nrows):
        row = []
        for y in range(len(columns)):
            row.append(str(sheet.cell_value(record, y)))
        cursor.execute("Insert into dim_incident (%s) Values ('%s');" % (",".join(names), "','".join(row)))

    # Close the cursor
    cursor.close()
    # Commit the transaction
    database.commit()
    # Close the database connection
    database.close()

    return
def SLA_Counter(file):
    '''this function creates a SLA counter in column number 22 which is incremented for every transfer'''

    wb = xlwt.Workbook()
    outsheet = wb.add_sheet('Page 1', cell_overwrite_ok=True)
    open = xlrd.open_workbook(file)
    insheet = open.sheet_by_name("Page 1")
    counter=0
    lis=[]
    dictn={}
    outsheet.write(0, 21, "SLA_Counter")
    for row in range(0, insheet.nrows):
        for col in range(0, insheet.ncols):
            outsheet.write(row, col, insheet.cell_value(row, col))
    for row in range(1, ((insheet.nrows) - 1)):

        if insheet.cell_value(row, 0) == insheet.cell_value(row + 1, 0) and insheet.cell_value(row,3) != "Resolved":
            if insheet.cell_value(row,3) == "Assignment Group":
                counter+=1
                outsheet.write(row, 21, counter)
                lis.append(counter)
            elif insheet.cell_value(row,3) == "Assigned to Duration":
                outsheet.write(row, 21, counter)

        if insheet.cell_value(row, 0) != insheet.cell_value(row + 1, 0) and insheet.cell_value(row, 3) != "Resolved":
            if insheet.cell_value(row,3) == "Assignment Group":
                counter+=1
                outsheet.write(row, 21, counter)
                lis.append(counter)
            if insheet.cell_value(row, 3) == "Assigned to Duration":
                outsheet.write(row, 21, counter)


        if insheet.cell_value(row, 0) != insheet.cell_value(row + 1, 0):
            counter = 0
            dictn[insheet.cell_value(row, 0)]=max(lis)
            lis=[]

    for x in range(1,insheet.nrows):
        if insheet.cell_value(x, 3) == "Resolved":
            outsheet.write(x, 21, (dictn[insheet.cell_value(x, 0)]+1))
            dictn[insheet.cell_value(x, 0)]+=1

    wb.save(file)
    return
