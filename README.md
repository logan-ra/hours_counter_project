# hours_counter_project

I work at a public library in Northern Kentucky, and one of the duties of my department is to staff the three public desks and drive-thru window. Nearly everyone in our library works at the public desks, though for varying amounts of time. Some staff are hired exclusively for this purpose, while others fill in where needed.

All staff members access and share a Microsoft Excel workbook that shows where they're assigned during any given hour. Staff members can tell which desk they're assigned to by finding their initials on that day's worksheet.

The purpose of my project is to iterate over a range of cells that encompasses each day's hourly desk assignments, count how many times each person's initials appear, and log that number on a worksheet at the end of the workbook. That number represents the amount of hours that staff member spent on desk during that week. The scheduler can then use this information to ensure everyone is working the appropriate amount of hours on desk.

**The three features included in my assignment are:**

##### Category 1:

**Create a disctionary or list, populate it with several values, retrieve at least one value, and use it in your program.**

##### Category 2:

**Read data from an external file, such as text, JSON, CSV, etc, and use that data in your application.**

##### Category 3:

**Display data in tabular form.**


To use this project, install the [openpyxl](https://openpyxl.readthedocs.io/en/stable/#) library. Run the file "main.py," which will prompt you to enter the file path where you've saved the Excel workbooks contained in the "Desk schedules files" folder.

If everything goes according to plan, each Excel workbook should be updated with a new column on the worksheet "Staff Initials."