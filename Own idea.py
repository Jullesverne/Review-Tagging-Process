# https://openpyxl.readthedocs.io/en/stable/tutorial.html library being used
from unittest.mock import NonCallableMagicMock
import pandas as pd 
import numpy as np
import re 
import nltk 
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
from openpyxl import load_workbook as lw 
import copy


nltk.download('stopwords') # CONSIDER CHANGING WHAT THESE ARE, MAYBE MAKE MY OWN


# Before any more changes, make sure this is linear, and figure out why I have extra tags in the tag_do doc, then look at why num_tags_dic is fucked up, 
# it is recording way too many of some types of tags (product - positive) because
# that would also be throwing off my word count in masters, and potentially word count in tag dictionary too 
# basically just go through and revalidate each step
# also re read through everything and make more functions to clean up legibility / how variables are named 

# then rethink how I am tagging because lol is showing
# up way too much for how under represented it should actually be 
# score could be divided by number of instances of those tags? instead of word specificic? or maybe divide it by total tags - number of those? play
# playing around with scoring could definitely help, could change to divided by unique number of words for that type of tag, or change it to divided by number of that
# type of review, maybe better than number of unique or otherwise words or do it as number of 
# times words show up for that specific tag divided by number of times that tag occurs


# definitely I think make my own version of stopwords, also consider removing the stemming thing, or updating it
# to just take off ing, ly, ad verb endings bc horrible should be a strong indication of negative 
# and look at different success rates for different tags --> what do I do well on vs not well 
# can change 1) way scorig is made 2) what are stop words 3) how stemming is done (seems to lower success rate
# could change it to look at just most common words for a specific review type, and instead of a score, just increase by 1 if it has that word? 

# NOTE THAT I AM NOW NOT ADDING THE SCORE BUT INSTEAD ADDING A NUMBER JUST IF review has same word as one in dictionary -- SEEMED TO GREATLY IMPROVE UP TO 8%

# also re look at my accuray maker because I don't 100% trust it 

# could also make the score percent of instances of that word compared to average percent of that word for all reviews

# look at percents of different tags to see which one I am over/under reporting compared to historical tags


workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/better.xlsx') 
sheet = workbook.active

def review_cleaner(rev):
    review = re.sub('[^a-zA-Z]', ' ',rev)
    review=review.lower()
    review=review.split()
    ps=PorterStemmer()
    review = [ps.stem(word) for word in review if not word in set(stopwords.words('english'))] 
    review =' '.join(review)
    return review

def add_to_main(word, master):
    if word in master:
        master[word]+=1
    else:
        master.update({word:1})
    return master

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

def make_score_all(master, tag_dic):
    for tag in tag_dic.keys():
        for word in tag_dic[tag].keys():
           tag_dic[tag][word]= tag_dic[tag][word]/master[word]
    return tag_dic


def make_score_num_tag(tags, num_tags):
    for tag in tags.keys():
        for word in tags[tag].keys():
            tags[tag][word] = tags[tag][word]/num_tags[tag]
    return tags



def make_score_tag_specific(tag_dic): 
    for tag in tag_dic.keys():
        values=tag_dic[tag].values()
        total = sum(values)
        for word in tag_dic[tag].keys():
           tag_dic[tag][word]= tag_dic[tag][word]/total
    return tag_dic



def score_creator(review, score_dic):
    score = 0 
    rev_iter = review.split()
    for word in rev_iter:
        if word in score_dic.keys():
            score+=score_dic[word]
            #score+=1
    score= float(score)
    score = score % len(rev_iter) / 1.000000000000
    return score

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
        for word in clean_rev.split():
            add_to_main(word, master_dic)
            add_to_tag_dic(word, tag.value.lower(), tags_dic)
            add_to_main(tag.value.lower(), num_tags_dic) # I think there is an error here
    x+=1

print(num_tags_dic.items())

tags_score_num = copy.deepcopy(tags_dic)

tags_score_num = make_score_num_tag(tags_score_num, num_tags_dic)
# now have both dictionaries, setting up the score for them WILL NEED TO ANNOTATE THIS BETTER LATER
tags_score_specific = copy.deepcopy(tags_dic)

tags_score_specific = make_score_tag_specific(tags_score_specific)

tags_score_all = make_score_all(master_dic, tags_dic) #look at differing success rates for two different types of scoring


fresh = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx') 
reviews = fresh.active

x=1
rev_col = 'A' #col reviews are in
review = reviews[str(rev_col)+str(x)]
tag_col = 'C' # where you want tag placed
tag_true = 'B'
correct_tag = str(tag_true)+str(x)

RS = 0 
RA = 0 
WS = 0 
WA = 0 
while review.value != None:
    review_location = str(rev_col)+ str(x)
    tag_location = str(tag_col)+str(x)
    review = reviews[review_location]
    correct_tag = str(tag_true)+str(x)
    if isinstance(reviews[correct_tag].value, str):
        correct_tag_value = reviews[correct_tag].value.lower()
    review_scores_all = {}
    review_scores_specific = {}
    review_scores_num = {}
    if isinstance(review.value, str):
        cleaned= review_cleaner(review.value)
        if len(cleaned.split())>0:
            for tag in tags_score_all.keys():
                    rating_specific = score_creator(cleaned,tags_score_specific[tag])
                    review_scores_specific[tag] = rating_specific

                    rating_all = score_creator(cleaned,tags_score_all[tag])
                    review_scores_all[tag]=rating_all

                    review_num_tags=score_creator(cleaned, tags_score_num[tag])
                    review_scores_num[tag] = review_num_tags
            maxkey_specific = max(review_scores_specific, key=review_scores_specific.get)
            maxkey_all = max(review_scores_all, key=review_scores_all.get)
            maxkey_num = max(review_scores_num, key=review_scores_num.get)
    reviews[tag_location] = maxkey_specific



    # DOING THIS MAKE A NEW WAY TO LOOK AT WHICH TAGS ARE tagged more accurately vs not beyond just printing, actually count
    # prolly another dictionary
    if isinstance(correct_tag_value, str):
        if maxkey_specific == correct_tag_value.lower():
            RS+=1
        elif maxkey_specific != correct_tag_value.lower():
            WS+=1
            #print('my specific prediction was '+ maxkey_specific + ' and correct prediction was ')
            #print(correct_tag_value)
        elif maxkey_num == correct_tag_value.lower():
            RA +=1
        elif maxkey_num != correct_tag_value.lower():
            WA+=1
    else:
        print('old tag was weird')
        print(correct_tag_value)
        #print('my all prediction was ' + maxkey_all + ' and correct prediction was ')
        #print(correct_tag_value)
    
    #print('                                     ')
    #print(review_scores_all.items())
    #print(review_scores_specific.items())
    x+=1



print('I MADE IT OUT')
print(x)
total = RA+RS+WA+WS

print(total)
print('correct specific ' + str(RS))

print('correct num '+ str(RA))


fresh.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx') 