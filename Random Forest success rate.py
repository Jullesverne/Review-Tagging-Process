# this file has the accuracy rates for your test based off your data and tweeks to it can give you a better sense of 
# the accuracy, does not put reviews into excel sheet, just shows accuracy!!

# importing everything thats needed 
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

# this takes your prediction and compares it to the tagged data, does AB testing
# will make a new one when not just doing AB testing
def accuracy_percent_ab_tag(y_pred, y_test):
    false_pos=0
    false_neg=0
    true_pos=0
    true_neg=0
    i=0
    while i<len(y_pred):
        if y_pred[i] == 1 and y_test[i]==1:
            true_pos+=1
        elif y_pred[i] == 0 and y_test[i]==0:
            true_neg+=1
        elif y_pred[i]== 1 and y_test[i] == 0:
            false_pos+=1
        elif y_pred[i] == 0 and y_pred[i] == 1:
            false_neg+=1
    i+=1
    total = false_neg+false_pos+true_neg+true_pos
    print('percent correct pos '+ str(true_pos/total*100))
    print('percent correct neg '+ str(true_neg/total*100))
    print('percent false pos '+ str(false_pos/total*100))
    print('percent false neg '+ str(false_neg/total*100))
    print('total reviews tagged '+ str(total))


# you need to pass this the array of the two columns you want to compare from the sheet
# may also need to change this if the values are different, i.e one and zero vs strings
def accuracy_percent_off_columns(y_pred, y_test):
    correct = 0 
    wrong = 0 
    i=0
    while i<len(y_pred):
        if y_pred[i] == y_test[i]:
            correct+=1
        else:
            wrong +=1
    i+=1
    total = correct+wrong
    print('percent correct '+ str(correct/total*100))
    print('percent wrong '+ str(wrong/total*100))

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


# now splitting into the training and test set, changing test size 
# could improve results
X_train, X_test, y_train, y_test = train_test_split(X, Y, test_size = 0.2)

for i,x in enumerate(y_train):
    y_train[i]=int(x)
    

# playing around with number of n_estimators can change accuracy 
# don't know what entropy is tbh
model = RandomForestClassifier(n_estimators=501, criterion='entropy')

model.fit(X_train, y_train)

# have build our model, now run it on the test data

y_pred=model.predict(X_test)











