#print('top of file')
import pandas as pd
file_name =  'c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx'

import pandas as pd
df = pd.read_excel(io=file_name)
print(df.head(5))  # print first 5 rows of the dataframe
#import pathlib
#pathlib.Path().resolve()