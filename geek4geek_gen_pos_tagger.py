# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used

# THIS TOOL CAN BE USED TO TAG ONE SPECIFIC TYPE, I.E A-B TAGGING NOT THIRD OPTIONS
# totally flexible on what you choose to tag, just need to prep it in training set 
# WHICH I WILL WRITE DOCUMENTATION FOR 
from tracemalloc import stop
import pandas as pd 
import numpy as np
import re 
import nltk 
from sklearn.feature_extraction.text import CountVectorizer
nltk.download('stopwords') 

# to remove stopwords (super common ones)
from nltk.corpus import stopwords

# for Stemming propose - just get roots of words no -ly -ing
from nltk.stem.porter import PorterStemmer

# this is for excel processing
from openpyxl import load_workbook as lw 

# you need to put in your local path here
workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Gdoc_rev_prep.xlsx') 
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

# you need to drop in header for review and tag column here
review_column = 'D' 
tag_column = 'C' 

x=1 
cell = sheet[str(review_column)+str(x)]

corpus= []
#rn this loop is taking dirty review and cleaning it, and append to corpus
# to create the training set, set the max value for x (while x<=...) 
while x>=1:
    review_location = str(review_column)+ str(x)
    cell = sheet[review_location]
    if cell.value == None:
        break # this it for when we get to end of list
    elif isinstance(cell.value, str):
        corpus.append(review_cleaner(cell.value))
    x+=1

# now creating a bag of words

cv = CountVectorizer()
# X contains corpus, dependent variable
X=cv.fit_transform(corpus).toarray()


# Y contains answers, if review pos or neg, need .value to actually access 1 or 0 
Y = sheet['E'] 

# now splitting into the training and test set






#workbook.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/updated.xlsx') # you put in the file path and name of where you want 









