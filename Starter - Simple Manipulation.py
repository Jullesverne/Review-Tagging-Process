#https://openpyxl.readthedocs.io/en/stable/tutorial.html library I'm using
# HARD CODE the column reviews are in 
# then can interate down them numerically
# and print in the new data to the column on the right (also hard coded)

import pandas as pd
from openpyxl import load_workbook as lw
file_name =  'c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx'

workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx')
sheet= workbook.active
sheet["A1"] = 'changer'
c=sheet['A2']
workbook.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx')



# HARD CODE the column reviews are in 
# then can interate down them numerically
# and print in the new data to the column on the right (also hard coded with number and name)


# this is just a simple iterator, should be duplicate once
# I have copied the data from one column to next
for row in sheet.values:
    for value in row:
        print(value)

