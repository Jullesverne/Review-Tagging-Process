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
nltk.download('stopwords') # CONSIDER CHANGING WHAT THESE ARE, MAYBE MAKE MY OWN


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

# similar to above, but supposde to make the value the number of times that word was seen for that tag divided by 
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

#loading in data to build the model 
workbook = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/baby.xlsx') 
sheet = workbook.active

# need to update these with correct columns, this is taking in the already tagged examples and building the model 
review_column = 'A' 
tag_column = 'B' 
x=1 
cell = sheet[str(review_column)+str(x)]
master_dic = {}
tags_dic = {}
num_tags_dic = {} 

# this creates the master dic with total word count, and the 
# tag dic with cound of word specific to that review
# BOth of which need to be verified, and also need to add in the num_tags_dic to actually work
while cell.value!= None:
    review_location = str(review_column)+ str(x)
    pre_tagged_value = str(tag_column)+str(x)
    cell = sheet[review_location]
    tag=sheet[pre_tagged_value]
    if isinstance(cell.value, str) and isinstance(tag.value, str):
        clean_rev = review_cleaner(cell.value)
        add_to_main(tag.value.lower(), num_tags_dic)
        for word in clean_rev.split():
            add_to_main(word, master_dic)
            add_to_tag_dic(word, tag.value.lower(), tags_dic)
    x+=1

doc = aw.Document()
builder=aw.DocumentBuilder(doc)

#this is copying the tags_dic so I can have different scoring dictionaries 
# NEED TO VERIFY EACH OF THE SCORING DICTIONARIES AND FUNCITONS AS WELL
tags_score_num = copy.deepcopy(tags_dic)
tags_score_num = make_score_num_tag(tags_score_num, num_tags_dic)

builder.write('      score for num tags - should be word count divided by number of that type of tag')
for tag in tags_score_num.keys():
    builder.write(tag)
    builder.write(str(tags_score_num[tag].items()))

tags_score_specific = copy.deepcopy(tags_dic)
tags_score_specific = make_score_tag_specific(tags_score_specific)

builder.write('    score for specific - should be word count divded by total words used for that type of tag ')
for tag in tags_score_specific.keys():
    builder.write(tag)
    builder.write(str(tags_score_specific[tag].items()))

tags_score_all = copy.deepcopy(tags_dic)
tags_score_all = make_score_all(master_dic, tags_score_all) 

builder.write('     score for all - should be word count divided by total number of times that word was used for all tags ')
for tag in tags_score_all.keys():
    builder.write(tag)
    builder.write(str(tags_score_all[tag].items()))

doc.save("c:/Users/HP/Desktop/Review-Tagging-Process/reader.docx")

# now loading in new reviews that I am going to generate tags for 
fresh = lw(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx') 
reviews = fresh.active

x=1
rev_col = 'A' #col reviews are in
review = reviews[str(rev_col)+str(x)]
tag_col = 'C' # where you want tag placed
tag_true = 'B' # because I am using a test set this is where the true tags are actually contained
correct_tag = str(tag_true)+str(x)

RS = 0 
RA = 0 
WS = 0 
WA = 0 

when_write = {}
when_wrong = {}
while review.value != None:
    review_location = str(rev_col)+ str(x)
    review = reviews[review_location]
    tag_location = str(tag_col)+str(x) # where its gonna be put
    correct_tag = str(tag_true)+str(x) # one I am checking against
    # this if statement is just for comparisons sake 
    if isinstance(reviews[correct_tag].value, str):
        correct_tag_value = reviews[correct_tag].value.lower()

    review_scores_all = {}
    review_scores_specific = {}
    review_scores_num = {}

    #now we are generating the score for each potential tag of a review 
    if isinstance(review.value, str):
        #print('printing review')
        #print(review.value)
        cleaned = review_cleaner(review.value)
        if len(cleaned.split())>0:
            for tag in tags_score_all.keys():
                    #print('printing tag')
                    #print(tag)
                    # verify that this section works correctly 
                    rating_specific = review_score_creator(cleaned,tags_score_specific[tag])
                    review_scores_specific[tag] = rating_specific
                    #print('score for above review from specific ')
                    #print(rating_specific)

                    rating_all = review_score_creator(cleaned,tags_score_all[tag])
                    review_scores_all[tag]=rating_all
                    #print('score for above from all')
                    #print(rating_all)

                    # I think next two lines could have an error
                    review_num_tags= review_score_creator(cleaned, tags_score_num[tag])
                    review_scores_num[tag] = review_num_tags
                    #print('score for above from num_tags')
                    #print(review_num_tags)

            # these three lines are finding the tag with the highest score for the review
            #maxkey_specific = max(review_scores_specific, key=review_scores_specific.get)
            #maxkey_all = max(review_scores_all, key=review_scores_all.get)
            maxkey_num = max(review_scores_num, key=review_scores_num.get)

    # this is putting the tag into the location specified earlier
    reviews[tag_location] = maxkey_num
    
    

    # This whole segment is to test accuracy, will be rehauled before considered seriously 
    if isinstance(correct_tag_value, str):
        if maxkey_num == correct_tag_value.lower():
            RS+=1
            add_to_main(maxkey_num, when_write)
        elif maxkey_num != correct_tag_value.lower():
            WS+=1
            add_to_main(correct_tag_value.lower(), when_wrong)
        #elif maxkey_num == correct_tag_value.lower():
            #RA +=1
        #elif maxkey_num != correct_tag_value.lower():
            #WA+=1
    else:
        print('old tag was weird')
        print(correct_tag_value)
    # increase x to iterate to the next tag
    x+=1



# this is all just accuracy testing
print('when correct')
print(when_write.items())

print('when wrong')
print(when_wrong.items())

total = RA+RS+WA+WS

print('number write using new score')
print(RS/total*100)

print('total')
print(total)
# this is saving the now tagged reviews to a file you specify
fresh.save(filename='c:/Users/HP/Desktop/Review-Tagging-Process/to_tag.xlsx') 