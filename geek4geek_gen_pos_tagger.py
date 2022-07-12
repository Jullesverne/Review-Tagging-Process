# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used

# THIS TOOL CAN BE USED TO TAG ONE SPECIFIC TYPE, I.E A-B TAGGING NOT THIRD OPTIONS
# totally flexible on what you choose to tag, just need to prep it in training set 
# WHICH I WILL WRITE DOCUMENTATION FOR - but can only tag one type 
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
tag_column = 'E' 

x=1 
cell = sheet[str(review_column)+str(x)]
length_cor=1
corpus= []
Y=[]
#rn this loop is taking dirty review and cleaning it, and append to corpus
# to create the training set, set the max value for x (while x<=...) 
while cell.value != None:
    review_location = str(review_column)+ str(x)
    pre_tagged_value = str(tag_column)+str(x)
    cell = sheet[review_location]
    tag=sheet[pre_tagged_value]
    if isinstance(cell.value, str) and tag.value != None:
        corpus.append(review_cleaner(cell.value))
        Y.append(tag.value)
    #adding just number onto corpus so X and Y have same length
    elif tag.value!= None:
        corpus.append(str(cell.value))
        Y.append(tag.value)
    x+=1
    length_cor+=1

# now creating a bag of words

cv = CountVectorizer()
# X contains corpus, dependent variable
X=cv.fit_transform(corpus).toarray()


# Y contains answers, if review pos or neg, need .value to actually access 1 or 0 
# Y = sheet['E']

# now splitting into the training and test set, changing test size 
# could improve results
print('about to make test/train set')


#X_train, X_test, y_train, y_test = train_test_split(X, Y, test_size = 0)
#y_train=np.ndarray.tolist(y_train)
# NEED THIS LINE BELOW OR GET VARIABLE TYPE UNKNOWN ERROR
#print('checking y_train data type')
#print(type(y_train))

#for i,x in enumerate(y_train):
    #y_train[i]=int(x)
    



# playing around with number of n_estimators will 
# make more or less accurate

#entropy criterion is fancy math stuff, not 100% sure how it works
# bc also have gini like gini coeffecient 
model = RandomForestClassifier(n_estimators=501, criterion='entropy')

#print('y type checking again')
#print(type(y_train))

model.fit(X, Y)

# RIGHT HERE is where you would put in the new reviews I think and get the corresponding tag, as y_pred would be list
# of how it should be tagged, and X_test would be the reviews I want tags generated for 

# LOAD IN NEW DATA FOR IT TO GET TESTED ON





y_pred=model.predict(new_taggers)



print(y_pred.shape)
print('above is prediction shape below is test shape')
print(len(y_test))

print('type of y_predictions')
print(y_pred[0])
print(y_test[0])


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

print(X_test[0])
#print(str(y_pred[0]) + ' ' + X_test[0])
print('percent correct pos '+ str(true_pos/total*100))
print('percent correct neg '+ str(true_neg/total*100))
print('percent false pos '+ str(false_pos/total*100))
print('percent false neg '+ str(false_neg/total*100))
print('total reviews tagged '+ str(total))



# now figure out way to get tag back to corresponding reivew --> will likely be easier when not using X_test but a new file 
# so try that route to start with 

# then run this on the other batch of reviews, using the whole first batch as the training set

# set it up to put predictions next to reviews so can see direct comparison, and to make sure y_pred keeps order with reviews

#workbook.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/updated.xlsx') # you put in the file path and name of where you want 









