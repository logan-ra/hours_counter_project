#Load workbook. Load the worksheet 'Staff Initials'
import openpyxl

def hours_counter(weekly_schedule):
    wb = openpyxl.load_workbook(weekly_schedule)
    staff_initials = wb['Staff Initials']

#Create empty list to insert values of cells
    initials = []

#rowOfCellOb == giant rectangle from F7:T29
#For each thing (in this case cell) in rowOfCellOb, find the actual content of the cell (using .value) and add it to the empty list "initials" (using .append)
    for sheet in wb.worksheets:
        for rowOfCellOb in sheet['F7':'T29']:
            for each in rowOfCellOb:
            #Avoid adding empty cells to list with "if". If the value of the cell is not None, then you can append it to the list.
                if each.value is not None:
                    initials.append(each.value)
#To compensate for too much data on Sunday, Friday, and Saturday, re-populate the "initials" list. Take entry ("each") for each item in "initials" with a length less than or equal to 4. This leaves two, three, and four letter intials. 
    initials = [each for each in initials if len(each) <= 4]

#Create an empty dictionary called "hours". We will make each item in "initials" a key, and the value will be a count of the amount of instances of that key.
    hours = {}
    for i in initials:
        if i in hours:
            hours[i] += 1
        else:
            hours[i] = 1
    print(hours)
#Starting at row 2 to skip headers. For rows from row 2 to the end of the used spreadsheet, do the following:
    for rowNum in range(2, staff_initials.max_row + 1):
    #create a variable called "enter". Enter is the value of each cell in the first column of "Staff Initials"
        enter = staff_initials.cell(row=rowNum, column = 1).value
    #If the cell values in enter are in "hours" then starting at column 4, input those values on the "Staff Initials" worksheet with "enter" as the key
        if enter in hours:
            staff_initials.cell(row=rowNum, column=4).value = hours[enter]
        else:
            staff_initials.cell(row=rowNum, column=4).value= 0

    wb.save(weekly_schedule)