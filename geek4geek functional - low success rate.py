# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used
# low success rate, can be used to tag one specific tag (A-B) but will not generate tags for more than 1 thing
# to change what it is tagging, you need to update the data model is trained on, currently trained to 
#tag general - positive


# neccesary imports 
from ctypes import sizeof
from tracemalloc import stop
import pandas as pd 
import numpy as np
import re 
import nltk 
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import confusion_matrix
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
from openpyxl import load_workbook as lw 
nltk.download('stopwords') 

# this section is building your model, you would need to put in your local path to the data you want to train 
# the model on into the filename section 
# make the tag you want 1 and all others 0 for what you are training it on, and the model will tag 
# the new reviews with a 1 if it falls into the category you want or a 0 if not

workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/Gdoc_rev_prep.xlsx') 
sheet = workbook.active

#this function cleans the reviews, getting rid of extra words and endings
def review_cleaner(rev):
    review = re.sub('[^a-zA-Z]', ' ',rev)
    review=review.lower()
    review=review.split()
    ps=PorterStemmer()
    review = [ps.stem(word) for word in review if not word in set(stopwords.words('english'))] 
    review =' '.join(review)
    return review

# Review Column and Tag Column need the letter for their respective columns on data you are training from 
review_column = 'D' 
tag_column = 'E' 

x=1 
cell = sheet[str(review_column)+str(x)]
corpus= []
Y=[]

# this loop is cleaning reviews and making a training set 
while cell.value != None:
    review_location = str(review_column)+ str(x)
    pre_tagged_value = str(tag_column)+str(x)
    cell = sheet[review_location]
    tag=sheet[pre_tagged_value]
    if isinstance(cell.value, str) and tag.value != None:
        corpus.append(review_cleaner(cell.value))
        Y.append(tag.value)
    elif tag.value!= None:
        corpus.append(str(cell.value))
        Y.append(tag.value)
    x+=1


# Max feature is important, sets the number of variables you will consider, playing with this could 
# change accuracy, also note may have issue if too large
cv = CountVectorizer(max_features = 2000)

X=cv.fit_transform(corpus).toarray()

model = RandomForestClassifier(n_estimators=501, criterion='entropy')
model.fit(X, Y)

# now we have built our model, below is using it to classify new reviews
# here you need to put in your local file path to the new reviews you want tagged
new_rev=lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx')
reviews = new_rev.active

# this should be the column that contains reviews 
review_column = 'A'
x=1
rev2= reviews[str(review_column)+str(x)]
body= []

while rev2.value != None:
    review_location = str(review_column)+str(x)
    rev2 = reviews[str(review_column)+str(x)]
    if isinstance(rev2.value, str):
        body.append(review_cleaner(rev2.value))
    else:
        body.append(str(rev2.value))
    x+=1
body2=cv.fit_transform(body).toarray()

#generating predictions
y_pred=model.predict(body2)

#tag column needs letter of column you want the tags to be put into 
tag_column = 'C'

iter=1
index=0
while index<len(y_pred):
    tag_location = str(tag_column)+str(iter)
    if y_pred[index] == 1:
        reviews[tag_location] = 'General - positive'
    else:
        reviews[tag_location] = 'Other'
    iter+=1
    index+=1

#you need to specify file path that tagged reviews should be saved to, this is for your local computer
new_rev.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/pleasee.xlsx') 