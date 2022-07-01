import pandas as pd
from openpyxl import load_workbook as lw
print('working')
file_name =  'c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx'

workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx')
sheet= workbook.active
sheet["A1"] = 'changer'
c=sheet['A2']
print(c)
workbook.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx')

#df = pd.read_excel(io=file_name)
#print(df.head(5))  # print first 5 rows of the dataframe
#import pathlib
#pathlib.Path().resolve()