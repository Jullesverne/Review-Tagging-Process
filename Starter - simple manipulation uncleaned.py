# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used

#importing things to open, clean and tag, THESE COULD DEPRECEATE
from tracemalloc import stop
import pandas as pd 
import numpy as np
import re 
import nltk 
nltk.download('stopwords') 

# to remove stopwords (super common ones)
from nltk.corpus import stopwords

# for Stemming propose - just get roots of words no -ly -ing
from nltk.stem.porter import PorterStemmer

from openpyxl import load_workbook as lw # this is for excel processing

workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Sample.xlsx') # you need to put in your local path here
sheet = workbook.active

#function to clean review - needed my route or geek4geek route - NOT for github guys route
def review_cleaner(rev):
    review = re.sub('[^a-zA-Z]', ' ',rev)
    review=review.lower()
    review=review.split()
    #create object to take main stem of each word
    ps=PorterStemmer()
    #loop through each word to stem and cut off the stopwords (i.e the a an etc) 
    # COULD BE ISSUE HERE WITH STOPWORDS BEING THEIRS AND NOT MINE -- IF BAD ACCURACY LOOK INTO MAKING OWN SET OF STOPWORDS
    review = [ps.stem(word) for word in review if not word in set(stopwords.words('english'))] 
    review =' '.join(review)
    return review

review_column = 'B' # you need to drop in the column that has reviews here
tag_column = 'C' # you need to drop in the column where you want to put tag here

x=1 # excel does not start iterating from 0, it starts iterating from 1 for rows, A for columns 

cell = sheet[str(review_column)+str(x)]

# up here need to go through pre tagged reviews and create scores (dictionary of dictionary) so 
# when we get to while loop all we'd be doing is calculating score for each possible tag 
# will need to open file of pre tagged reviews to generate dictionaries 

#rn this loop is taking dirty review and cleaning it
while x>=1 :
    review_location = str(review_column)+ str(x)
    tag_location = str(tag_column) + str(x)
    cell = sheet[review_location] # now cell variable contains content of Review 
    if cell.value == None:
        break # this it for when we get to end of list

    # need a function here to generate the tag for the review --> this will be the iterating through dictionaries 

    sheet[tag_location] = review_cleaner(cell.value) # instead of cell.value this will be the tag that review gets
    #print(cell.value)
    x+=1


workbook.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/updated.xlsx') # you put in the file path and name of where you want 
# the updated document saved









