import openpyxl
wb = openpyxl.load_workbook('13 - 19 February 2022.xlsx')
sheet = wb['Sun13February']

initials = []

for rowOfCellOb in sheet['F7':'O13']:
  for each in rowOfCellOb:
    initials.append(each.value)
    print(initials)

#Use RegEx to capture cell .value with length of 2,3,or 4 characters?
#Could use function to say "if .value of cell contains str<=len(4)" then .append elif cell .value =/= azAZ ignore?