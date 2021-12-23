'''
Project to automate a task for the Transmission Company of Nigeria (TCN) control room.
The aim of this project is to automatically convert numbers in an excel file into
proper time formats. For example:
1100 -> 11:00
1234 -> 12:34
725  -> 07:25
etc.
'''
import time
import xlrd
import xlwt
from xlutils.copy import copy

#Function to convert number to time
def NumberToTime(number):
    if type(number) == float:
        number = int(number)
    if type(number) != int:
        return None
    numToStr = str(number)
    length = len(numToStr)

    if length == 1:
        timeFormat = ['0','0',':','0',numToStr]
    elif length == 2:
        timeFormat = ['0','0',':',numToStr[0],numToStr[1]]
    elif length == 3:
        timeFormat = ['0',numToStr[0],':',numToStr[1],numToStr[2]]
    elif length == 4:
        timeFormat = [numToStr[0],numToStr[1],':',numToStr[2],numToStr[3]]
    else:
        return None

    return "".join(timeFormat)

#Function to get index of desired spreadsheet cell
def GetColumnIndex(column):
    cells = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    return cells.rfind(column) % 26

#Function to read contents of an entire column in a spreadsheet
def ReadColumn(file,columnIndex):
    global rowData
    global wb #excel workbook for reading
    global sheet
    rowData = []
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0,0)
    for rowIndex in range(1,sheet.nrows):
        rowData.append(sheet.cell_value(rowIndex,columnIndex))

#Function to modify contents of an entire column in a spreadsheet
def WriteColumn(file,columnIndex):
    #Making a writable copy of the excel workbook
    wbCopy = copy(wb)
    #First sheet to write to within the writable copy
    w_sheet = wbCopy.get_sheet(0)
    #Write Data
    for i in range(1,sheet.nrows):
        w_sheet.write(i,columnIndex,rowData[i-1])
    #Save workbook
    wbCopy.save(file)

#Main application
def Main():
    file = input("Enter filename/path: ")
    column = input("Enter column name (e.g. A,B,C) for time data: ")
    columnIndex = GetColumnIndex(column)
    ReadColumn(file,columnIndex)
    for integer in rowData:
        result = NumberToTime(integer)
        if result != None:
            rowData[rowData.index(integer)] = result
    WriteColumn(file,columnIndex)

while True:
    Main()
    print("File successfully updated\n")
    time.sleep(3)
    print("Do you want to exit?(y:yes, n:no)\n")
    response = input()
    if response in "yesYESYes":
        break

