#Load workbook and sheet
import openpyxl
wb = openpyxl.load_workbook('13 - 19 February 2022.xlsx')

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

print(initials)

#Create an empty dictionary called "hours". We will make each item in "initials" a key, and the value will be a count of the amount of instances of that key.
hours = {}
for i in initials:
    if i in hours:
        hours[i] += 1
    else:
        hours[i] = 1

print(hours)
