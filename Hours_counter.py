import openpyxl
wb = openpyxl.load_workbook('24 - 30 October 2021.xlsx')
Initials = []
sheet = wb['Sun24October']
for i in range(7,13):
  print (i, sheet.cell(row=i, column=6).value)
sheet['F7'].value