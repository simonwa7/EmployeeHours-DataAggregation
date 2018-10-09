import openpyxl
import datetime
from datetime import timedelta
import os

# "data" object layout (2D list):
# [[First, Last, Employee Number, Date, Hours]
# .
# .
# .
# [First, Last, Employee Number, Date, Hours]]
# Note: Hours worked only counts if job name was a server

DATA = []
COMPILED = []

def printFull():
    textFile = open("fullData.txt", "a")
    for row in DATA:
        for column in row:
            textFile.write(str(column))
            textFile.write(",")
        textFile.write("\n")
    textFile.close()

def printCompiled():
    textFile = open("compiledData.txt", "a")
    for row in COMPILED:
        for column in row:
            textFile.write(str(column))
            textFile.write(",")
        textFile.write("\n")
    textFile.close()

def getColumnMatches(ws):
    # first, last, id, shift_start/date, shift_end (if applicable), jobtype,
    # punch_type (if applicable), hours (if applicable)
    columns = [-1, -1, -1, -1, -1, -1, -1, -1]
    for index, col in enumerate(ws.iter_cols(min_row=1, max_col=20, max_row=1)):
        for cell in col:
            if((cell.value == "First Name") or (cell.value == "FirstName") or (cell.value == "first_name")):
                columns[0] = index
            if((cell.value == "Last Name") or (cell.value == "LastName") or (cell.value == "last_name")):
                columns[1] = index
            if((cell.value == "Login Name") or (cell.value == "EmployeeNumber") or (cell.value == "login_name") or (cell.value == "Employee Id")):
                columns[2] = index
            if((cell.value == "ShiftStartTime") or (cell.value == "ClockIn") or (cell.value == "Shift Start Time")):
                columns[3] = index
            if((cell.value == "Job Name") or (cell.value == "JobName") or (cell.value == "Job") or (cell.value == "name")):
                columns[5] = index
            if((cell.value == "PunchTime") or (cell.value == "Punch Time")):
                columns[4] = index
            if((cell.value == "PunchType") or (cell.value == "Punch Type") or (cell.value == "Punch Type ID") or (cell.value == "Punch TypeID")):
                columns[6] = index
            if(cell.value == "Hours"):
                columns[7] = index

    return columns

def calculateHours(shift_start, shift_end, job_type, punch_type):
    if((punch_type == "Shift End") or (punch_type == "Shift Punch")):
        if((job_type == "Server") or (job_type == "Team Member - Server")):
            return(shift_end - shift_start)

    return 0

def getData(ws, columns):
    rows = iter(ws.rows)
    next(rows)
    for row in rows:
        dataRow = ["First", "Last", "Employee Number", "Date", -1]
        dataRow[0] = row[columns[0]].value
        dataRow[1] = row[columns[1]].value
        dataRow[2] = row[columns[2]].value
        dataRow[3] = row[columns[3]].value
        if(columns[7] != -1):
            if((row[columns[5]].value == "Server") or (row[columns[5]].value == "Team Member - Server")):
                dataRow[4] = row[columns[7]].value
            else:
                dataRow[4] = 0
        else:
            dataRow[4] = calculateHours(row[columns[3]].value,
                         row[columns[4]].value, row[columns[5]].value,
                         row[columns[6]].value)
        if("REDACTED" in dataRow):
            # print(dataRow)
            continue
        if("Redacted" in dataRow):
            # print(dataRow)
            continue
        if(not dataRow[0]):
            # print(dataRow)
            continue
        if(not dataRow[1]):
            # print(dataRow)
            continue
        if(not dataRow[2]):
            # print(dataRow)
            continue
        if(not dataRow[3]):
            # print(dataRow)
            continue
        dataRow[2] = int(dataRow[2])
        DATA.append(dataRow)

def getID(elem):
    return elem[2]

def getDate(elem):
    return elem[3]

def compileData():
    DATA.sort(key=getDate)
    DATA.sort(key=getID)

    index = 0
    while(index < len(DATA)):
        new_row = [DATA[index][0], DATA[index][1], DATA[index][2], DATA[index][4]]
        hours = 0

        while(index < len(DATA)):
            if(DATA[index][2] != new_row[2]):
                break
            if(DATA[index][4] != 0):
                if(isinstance(DATA[index][4], datetime.timedelta)):
                    hours += (DATA[index][4].total_seconds()/3600)
                else:
                    hours += DATA[index][4]
            index += 1

        new_row[3] = hours
        COMPILED.append(new_row)

def run_doc(wb_name):
    wb = openpyxl.load_workbook(wb_name)
    ws = wb.active
    columns = getColumnMatches(ws)
    # print(columns)
    getData(ws, columns)
    # print(DATA)

# # format1 - check
# run_doc('BE00001883.xlsx')
# # format2 - check
# run_doc('BE00001824.xlsx')
# # format3 - check
# run_doc('BE00001819.xlsx')
# # format4 - check
# run_doc('BE00001870.xlsx')
# # format5 - check
# run_doc('BE00001844.xlsx')
# # grouped format1
# run_doc('BE00003498.xlsx')
# grouped format2
# run_doc('BE00003499.xlsx')
# grouped format3
# run_doc('BE00003540.xlsx')
# grouped format4
# run_doc('BE00003541.xlsx')
for file in os.listdir('./Dropbox/Simon, William'):
    # print(file)
    # if(file == 'BE00001831.XLSX'):
    #     run_doc('./Dropbox/Simon, William/'+file)
    run_doc('./Dropbox/Simon, William/'+file)
compileData()
printFull()
printCompiled()
