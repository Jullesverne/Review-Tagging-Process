# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used


# YOU NEED TO PUT IN COLUMN RANGES TO START WITH 

# Rows start with 1, Columns start with A



import pandas as pd
from openpyxl import load_workbook as lw
file_name =  'c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx'

workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx') # you need to put in your local path here
sheet = workbook.active

review_column = 'B' # you need to drop in the column that has reviews here
tag_column = 'C' # you need to drop in the column where you want to put tag here

# WE COULD HAVE AN OPTIONAL SECONDARY TAG COLUMN WOULD JUST NEED TO ADD IN THE THIRD COLUMN NAME 


x=1 # excel does not start iterating from 0, it starts iterating from 1 for rows, A for columns 

cell = sheet[str(review_column)+str(x)]

while cell != None :
    review_location = str(review_column)+ str(x)
    tag_location = str(tag_column) + str(x)
    cell = sheet[review_location]
    if cell.value == None:
        break
    sheet[tag_location] = cell.value # this line would be where functions come in and the will update with generated tag, not the cell.value 
    print(cell.value)
    x+=1


workbook.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/updated.xlsx')






