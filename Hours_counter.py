#Load workbook and sheet
import openpyxl
wb = openpyxl.load_workbook('13 - 19 February 2022.xlsx')
sheet = wb['Sun13February']

#Create empty list to insert values of cells
initials = []

#rowOfCellOb == giant rectangle from F7:O13
#For each thing (in this case cell) in rowOfCellOb, find the actual content of the cell (using .value) and add it to the empty list "initials" (using .append)
for rowOfCellOb in sheet['F7':'O13']:
  for each in rowOfCellOb:
    #Avoid adding empty cells to list with "if". If the value of the cell is not None, then you can append it to the list.
    if each.value is not None:
      initials.append(each.value)
print(initials)

#Create an empty dictionary called "hours". We will make each item in "initials" a key, and the value will be a count of the amount 
hours = {}
for i in initials:
    if i in hours:
        hours[i] += 1
    else:
        hours[i] = 1
print(hours)


#Use RegEx to capture cell .value with length of 2,3,or 4 characters?
#Could use function to say "if .value of cell contains str<=len(4)" then .append elif cell .value =/= azAZ ignore?