# can be used to tag one specific tag (A-B) but will not generate tags for more than 1 thing
# to change what it is tagging, you need to update the data model is trained on, currently trained to 
# tag general - positive

# NOTE THAT YOU NEED TO SET THE max value for vectorizer to be the smaller number of unique words from the training
# and actually tagging data


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

workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/baby.xlsx') 
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


# you need to pass this the array of the two columns you want to compare from the sheet
# will need to update these, but still useful 
def accuracy_percent_off_columns(y_pred, y_test):
    correct = 0 
    wrong = 0 
    i=0
    while i<len(y_pred):
        if y_pred[i] == y_test[i].value:
            correct+=1
        else:
            wrong +=1
        i+=1
    total = correct+wrong
    print('percent correct '+ str(correct/total*100))
    print('percent wrong '+ str(wrong/total*100))


# Review Column and Tag Column need the letter for their respective columns on data you are training from 
# tag column needs to be the location of the tags in already tagged reviews 
# review column is just location of the reviews
review_column = 'A' 
tag_column = 'B' 

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
cv = CountVectorizer(max_features = 20)

X=cv.fit_transform(corpus).toarray()


model = RandomForestClassifier(n_estimators=501, criterion='entropy')
model.fit(X, Y)

# now we have built our model, below is using it to classify new reviews
# here you need to put in your local file path to the new reviews you want tagged
new_rev=lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/baby1.xlsx')
reviews = new_rev.active

# this should be the column that contains reviews 
review_column = 'A'
x=1
rev2= reviews[str(review_column)+str(x)]
body= []

#print('error is around here, body2 has too many variables line 95 or something')
while rev2.value != None:
    review_location = str(review_column)+str(x)
    rev2 = reviews[str(review_column)+str(x)]
    if rev2.value == None:
        break
    if isinstance(rev2.value, str):
        body.append(review_cleaner(rev2.value))
    else:
        body.append(str(rev2.value))
    x+=1

cv = CountVectorizer(max_features = 20)
body2=cv.fit_transform(body).toarray()
# THERE IS SOME ERROR RIGHT HERE, BODY2 HAS TOO MANY THINGS, SHOULD ONLY HAVE 2 NOT THREE



#generating predictions
y_pred=model.predict(body2)

#tag column needs letter of column you want the tags to be put into 
tag_column = 'C'

iter=1
index=0

col_correct = reviews['B']

#accuracy_percent_off_columns(y_pred, col_correct)
while index<len(y_pred):
    tag_location = str(tag_column)+str(iter)
    if y_pred[index] == 1:
        reviews[tag_location] = 'General - positive'
    else:
        reviews[tag_location] = 'Other'
    iter+=1
    index+=1




#you need to specify file path that tagged reviews should be saved to, this is for your local computer
new_rev.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/baby1.xlsx') 
