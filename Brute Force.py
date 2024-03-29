# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used

# importing stuff, not all used tbh
from unittest.mock import NonCallableMagicMock
import pandas as pd 
import numpy as np
import re 
import nltk 
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
from openpyxl import load_workbook as lw 
import copy
import aspose.words as aw
nltk.download('stopwords')


#NOTE all the tag_num_dic stuff is currently commented out, will need to comment it back in when done


#takes a review as as string, cleans it and returns it as a list 
def review_cleaner(rev):
    review = re.sub('[^a-zA-Z]', ' ',rev)
    review=review.lower()
    review=review.split()
    ps=PorterStemmer()
    review = [ps.stem(word) for word in review if not word in set(stopwords.words('english'))] 
    review =' '.join(review)
    return review

#takes a word and a dictionary, puts word as key, and value is numebr of times that word has been seen 
def add_to_main(word, master):
    if word in master:
        master[word]+=1
    else:
        master.update({word:1})
    return master

# takes a word, tag and dictionary of dictionaries
# tag is key for high level dicitonary, word is key for lower level
# keeps track of word count for each specific tag
def add_to_tag_dic(word, tag, nest):
    if tag in nest:
        if word in nest[tag]:
            nest[tag][word]+=1
        else:
            nest[tag].update({word:1})
    else:
        nest.update({tag:{}})
        nest[tag].update({word:1})
    return nest

# takes two dictionaries, one with count of every word, other dicitonary of dictionary, and updates values
# in the dictioary that holds the word count for each tag to be the word count for each tag / total number of times that word
# was seen
def make_score_all(master, tag_dic):
    for tag in tag_dic.keys():
        for word in tag_dic[tag].keys():
           tag_dic[tag][word]= tag_dic[tag][word]/master[word]
    return tag_dic

# similar to above, but supposed to make the value the number of times that word was seen for that tag divided by 
# the number of reviews with that tag
def make_score_num_tag(tags, num_tags):
    for tag in tags.keys():
        for word in tags[tag].keys():
            tags[tag][word] = tags[tag][word]/num_tags[tag]
    return tags


# same as make_score_all, but instead of dividing by the total word count for a specific word
# divides by the total number of words used for that type of tag
def make_score_tag_specific(tag_dic): 
    for tag in tag_dic.keys():
        values=tag_dic[tag].values()
        total = sum(values)
        for word in tag_dic[tag].keys():
           tag_dic[tag][word]= tag_dic[tag][word]/total
    return tag_dic


# takes a review as a string and a dictionary specific to a tag (i.e key is word in that tag, value is that words score)
# and returns the reviews score for that type of tag
def review_score_creator(review, score_dic):
    score = 0 
    rev_iter = review.split()
    for word in rev_iter:
        if word in score_dic.keys():
            score+=score_dic[word]
            #score+=1
    score= float(score)
    score = score / len(rev_iter) / 1.000000000000
    return score

def accuracy(tag, correct):
    if tag.lower() == correct.lower():
        return 1
    else:
        return -1

def other_avg(tag, word, tag_dic, num_tags):

        word_count = 0 
        tag_count = 0 
        for tags in tag_dic.keys():
            if tags!= tag:
                tag_count+=num_tags[tags]
             
                if word in tag_dic[tags].keys():
                    word_count+=tag_dic[tags][word]
                    
        return word_count / tag_count


def make_score_nums_minus_avg(tags, num_tags):
    c1 = copy.deepcopy(tags)
    c2 = copy.deepcopy(num_tags)
    for tag in tags.keys():
        for word in tags[tag].keys():
            othes = other_avg(tag, word, c1, c2)
           
            tags[tag][word] = tags[tag][word]/num_tags[tag] - othes
    return tags

#loading in data to build the model THIS is specific to you
workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/better.xlsx') 
sheet = workbook.active

# need to update these with your column headers, this is taking in the already tagged examples and building the model 
review_column = 'G' 
tag_column = 'J' 
x=1 
cell = sheet[str(review_column)+str(x)]
master_dic = {}
tags_dic = {}
num_tags_dic = {} 

# this creates the master dic with total word count, and the 
# tag dic with cound of word specific to that review
while cell.value!= None:
    review_location = str(review_column)+ str(x)
    pre_tagged_value = str(tag_column)+str(x)
    cell = sheet[review_location]
    tag=sheet[pre_tagged_value]
    if isinstance(cell.value, str) and isinstance(tag.value, str):
        clean_rev = review_cleaner(cell.value)
        tag.value = tag.value.strip()
        add_to_main(tag.value.lower(), num_tags_dic)
        for word in clean_rev.split():
            add_to_main(word, master_dic)
            add_to_tag_dic(word, tag.value.lower(), tags_dic)
    x+=1

doc = aw.Document()
builder=aw.DocumentBuilder(doc)

#this is copying the tags_dic so I can have different scoring dictionaries 

tags_score_num = copy.deepcopy(tags_dic)
tags_score_num = make_score_num_tag(tags_score_num, num_tags_dic)


tags_score_specific = copy.deepcopy(tags_dic)
tags_score_specific = make_score_tag_specific(tags_score_specific)



tags_score_all = copy.deepcopy(tags_dic)
tags_score_all = make_score_all(master_dic, tags_score_all) 


tags_num_minus_others = copy.deepcopy(tags_dic)
tags_num_minus_others = make_score_nums_minus_avg(tags_num_minus_others, num_tags_dic)


# now loading in new reviews that I am going to generate tags for 
fresh = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx') 
reviews = fresh.active

x=1
rev_col = 'A' #col reviews are in
review = reviews[str(rev_col)+str(x)]
tag_col = 'C' # where you want tag placed
tag_true = 'B' # because I am using a test set this is where the true tags are actually contained
correct_tag = str(tag_true)+str(x)

RA = 0
RS = 0 
RN = 0 
RF = 0 

when_right = {}
when_wrong = {}
while review.value!= None:
    review_location = str(rev_col)+ str(x)
    review = reviews[review_location]
    tag_location = str(tag_col)+str(x) # where its gonna be put
    correct_tag = str(tag_true)+str(x) # one I am checking against
    copier = str('D')+str(x)

    if review.value == None:
        break
    # this if statement is just for comparisons sake 
    if isinstance(reviews[correct_tag].value, str):
        correct_tag_value = reviews[correct_tag].value.lower()
        correct_tag_value = correct_tag_value.strip()
    review_scores_all = {}
    review_scores_specific = {}
    review_scores_num = {}
    review_scores_minus = {}
    #now we are generating the score for each potential tag of a review 
    if isinstance(review.value, str):
        john = copy.deepcopy(review.value)
        cleaned = review_cleaner(review.value)
        if len(cleaned.split())>0:
            for tag in tags_score_all.keys():
                    

                    rating_specific = review_score_creator(cleaned,tags_score_specific[tag])
                    review_scores_specific[tag] = rating_specific
                    

                    rating_all = review_score_creator(cleaned,tags_score_all[tag])
                    review_scores_all[tag]=rating_all
                    

                    review_num_tags= review_score_creator(cleaned, tags_score_num[tag])
                    review_scores_num[tag] = review_num_tags

                    rating_minus_others = review_score_creator(cleaned, tags_num_minus_others[tag])
                    review_scores_minus[tag] = rating_minus_others
                    
            # these three lines are finding the tag with the highest score for the review
            maxkey_specific = max(review_scores_specific, key=review_scores_specific.get)
            maxkey_specific = maxkey_specific.strip()
            maxkey_all = max(review_scores_all, key=review_scores_all.get)
            maxkey_all = maxkey_all.strip()
            maxkey_num = max(review_scores_num, key=review_scores_num.get)
            maxkey_mins = max(review_scores_minus, key=review_scores_minus.get)
            if john != None:
                reviews[tag_location] = maxkey_num
                reviews[copier] = john
            maxkey_num = maxkey_num.strip()

    # this is putting the tag into the location specified earlier
    
    
    # This whole segment is to test accuracy
    if isinstance(correct_tag_value, str):
        larry = copy.deepcopy(correct_tag_value)
        joe = copy.deepcopy(correct_tag_value)
        bob = copy.deepcopy(correct_tag_value)
        RS += accuracy(maxkey_specific, larry)
        RA += accuracy(maxkey_all, correct_tag_value)
        RN += accuracy(maxkey_num, joe)
        RF += accuracy(maxkey_mins, bob)
    
    else:
        print('old tag was weird')
        print(correct_tag_value)
    # increase x to iterate to the next tag
    x+=1

builder.write('number correct for all scoring dic')
builder.write('     ')
builder.write(str(RA))

builder.write('  number correct for specific scoring dic')
builder.write('     ')
builder.write(str(RS))

builder.write('  number correct for num scoring dic')
builder.write('     ')
builder.write(str(RN))

builder.write('   number correct for fancy way')
builder.write(str(RF))
# this is all just accuracy testing

total = x-1
builder.write(  'total reviews were    ')
builder.write(str(total))

print('number write using new score')
print(RN/total)

print('total')
print(total)
# this is saving the now tagged reviews to a file you specify
fresh.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx') 
doc.save("c:/Users/HP/Desktop/Review-Tagging-Process/reader.docx")
