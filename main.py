import openpyxl
import os
from Hours_counter import hours_counter

path = 'C:/Users/lrains/Desktop/CodeKY/desk schedules'

obj = os.scandir(path)

for entry in obj:
    if entry.is_file():
        hours_counter(entry)