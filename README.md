# Automating Review Tagging
# The goal of this was to use a machine learning algorithim (in this class Random Forest) to automate the way in which reviews were tagged with their sentiments. 
# This allowed for a faster analysis of customer feedback, as the tagging process became nearly instantaneous at 15 seconds for 1300 reviews as opposed to the 8+ hours it would take to tag that many reviews by hand. 

# This tool could be used to tag any text you chooose, as it is not specific to reviews. The initial Random Forest 
# script was only set up to do binary tagging, and had a success rate in the mid to high 80's. 

# Seperately, after the Random Forest method was working, I decided to see how accurate brute force methods would be at tagging reviews in more than a binary, i.e ones that used no ML to tag at least 3 kinds of reviews.

# It turned out they were very bad, with a success rate in the mid to low teens. This experience helped grow my interest in ML and lead to me enrolling in a graduate course focused on deep neural nets. 
# All sensitive information (i.e customers reviews) has been cleaned from this Repo when I made it public.
