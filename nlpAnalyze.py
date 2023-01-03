# -----------------------------------------------------------------------------------------------------------------------------------------
# Packages
# -----------------------------------------------------------------------------------------------------------------------------------------
import pandas as pd
from openpyxl import load_workbook

import nltk
from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize, word_tokenize, RegexpTokenizer
import requests

from bs4 import BeautifulSoup
from urllib.request import Request, urlopen

# -----------------------------------------------------------------------------------------------------------------------------------------
# Importing given dictionaries.
# -----------------------------------------------------------------------------------------------------------------------------------------

cik_list = pd.read_excel("cik_list.xlsx")
LoMcDo_wordList = pd.ExcelFile("LoughranMcDonald_SentimentWordLists_2018.xlsx")
postve_wordList = pd.read_excel(LoMcDo_wordList, 'Positive', index_col = None, header = None)
negatve_wordList= pd.read_excel(LoMcDo_wordList, 'Negative', index_col = None, header = None)
constraining_dict = pd.read_excel("constraining_dictionary.xlsx")
uncertainity_dict = pd.read_excel("uncertainty_dictionary.xlsx")
op_data = pd.read_excel("Output Data Structure.xlsx")
stopwords_dict = open("StopWords_Generic.txt","r")
stop_wordsList = stopwords_dict.readlines()

# -----------------------------------------------------------------------------------------------------------------------------------------
# Variables 
# -----------------------------------------------------------------------------------------------------------------------------------------

word_count, secfname, positive_score, negative_score, polarity_score, average_sentence_length, percentage_of_complex_words, fog_index,\
complex_word_count, word_count, uncertainty_score, constraining_score, subjectivityScore, sentiment_score, positive_word_proportion, negative_word_proportion,\
uncertainty_word_proportion, constraining_word_proportion, constraining_words_whole_report = [[] for i in range(19)]

# -----------------------------------------------------------------------------------------------------------------------------------------
# Functions
# -----------------------------------------------------------------------------------------------------------------------------------------
def section1_1(words):
    pos_score, neg_score, uncertain_score, constrain_score, subjectivity_score = 0, 0, 0, 0, 0
    constrain_list = []
    for each in words:
        if each.upper() in postve_wordList.values:
            pos_score += 1
        if each.upper() in negatve_wordList.values:
            neg_score += 1
        if each.upper() in uncertainity_dict.values:
            uncertain_score += 1
        if each.upper() in constraining_dict.values:
            constrain_score += 1
            constrain_list.append(each)
    
    positive_score.append(pos_score)
    negative_score.append(neg_score)
    uncertainty_score.append(uncertain_score)
    constraining_score.append(constrain_score)
    
    subjectivity_score = (pos_score + neg_score) / (len(words) + 0.000001)
    subjectivityScore.append(subjectivity_score)
    
    pol_score = (pos_score - neg_score) / ((pos_score + neg_score) + 0.000001)
    polarity_score.append(pol_score)
    sentiment_score.append(sentimentScoreCategorization(pol_score))
    
    proportns(pos_score,neg_score,uncertain_score,constrain_score,words)
    
    constraining_words_whole_report.append(constrain_list)

#                                          -------------------------------------------------------------

def section2(words):
    
    avg_sent = len(words)/len(sentences)
    average_sentence_length.append(avg_sent)
    
    complex_count = section4(words)
    
    comp_percent = (complex_count/len(words)) * 100
    percentage_of_complex_words.append(comp_percent)
    
    
    fog_ind = 0.4 * (avg_sent + comp_percent)
    fog_index.append(fog_ind)
    
#                                          -------------------------------------------------------------

def section4(words):
    
    vowels = ['A','E','I','O','U','a','e','i','o','u']
    complex_excp = ['es', 'ed']
    check_complex = 0
    complex_count = 0

    for each in words:
        for i in range(len(each)):
            if each[i] in vowels:
                check_complex += 1

            if (i == len(each)-2) and (each[i]+each[i+1] in complex_excp):
                check_complex -= 1

        if check_complex > 2:
            complex_count += 1
            check_complex = 0

    complex_word_count.append(complex_count)
    return complex_count

#                                          -------------------------------------------------------------

def proportns(pos_score,neg_score,uncertain_score,constrain_score,words):
    pos_prop = pos_score/len(words)
    positive_word_proportion.append(pos_prop)

    neg_prop = neg_score/len(words)
    negative_word_proportion.append(neg_prop)

    uncertain_prop = uncertain_score/len(words)
    uncertainty_word_proportion.append(uncertain_prop)

    constrain_prop = constrain_score/len(words)
    constraining_word_proportion.append(constrain_prop)

#                                          -------------------------------------------------------------

def update():
    op_data['positive_score'] = positive_score
    op_data['negative_score'] = negative_score
    op_data['polarity_score'] = polarity_score
    op_data['sentiment_score'] = sentiment_score
    op_data['subjectivity_score'] = subjectivityScore
    op_data['average_sentence_length'] = average_sentence_length
    op_data['percentage_of_complex_words'] = percentage_of_complex_words
    op_data['fog_index'] = fog_index
    op_data['complex_word_count'] = complex_word_count
    op_data['word_count'] = word_count
    op_data['uncertainty_score'] = uncertainty_score
    op_data['constraining_score'] = constraining_score
    op_data['positive_word_proportion'] = positive_word_proportion
    op_data['negative_word_proportion'] = negative_word_proportion
    op_data['uncertainty_word_proportion'] = uncertainty_word_proportion
    op_data['constraining_word_proportion'] = constraining_word_proportion
    op_data['constraining_words_whole_report'] = constraining_words_whole_report

#                                          -------------------------------------------------------------

def sentimentScoreCategorization(polarityScore):
    if polarityScore < -0.5:
        return "Most Negative"
    elif polarityScore > -0.5 and polarityScore < 0:
        return "Negative"
    elif polarityScore == 0:
        return "Moderate"
    elif polarityScore > 0 and polarityScore < 0.5:
        return "Positive"
    elif polarityScore > 0.5:
        return "Very Positive"

#                                          -------------------------------------------------------------

def writeExcel():
    output = pd.ExcelWriter("Output Data Structure.xlsx", sheet = "Sheet1", engine = 'openpyxl')

    op_data.to_excel(output, index = False, startrow = 0)
    output.save()
    

# -----------------------------------------------------------------------------------------------------------------------------------------
# Main Code
# -----------------------------------------------------------------------------------------------------------------------------------------

 

indexCols = cik_list.loc[:,"CIK":"FORM"]
op_data = op_data.append(indexCols, sort = False)


make_link = "https://www.sec.gov/Archives/"

for i in range(cik_list.shape[0]):  # Making SECFNAME link
    link = make_link + cik_list["SECFNAME"][i]
    secfname.append(link)
    
op_data["SECFNAME"] = secfname



for i in range(op_data.shape[0]):
    print(i)
    requestUrl = op_data.SECFNAME[i]
    
    txt = requests.get(requestUrl) 

    req = Request(
        url=requestUrl, 
        headers={'User-Agent': 'XYZ/3.0'}
    )
    
    webpage = urlopen(req, timeout=10).read()
    txt = BeautifulSoup(webpage, "lxml").get_text(strip=True) # Cleaning for HTML codes
    
    string_encode = txt.encode("ascii", "ignore") # Cleaning ASCII Codes
    txt = string_encode.decode()
    
    sentences = sent_tokenize(txt)
    tokenizer = RegexpTokenizer("[\w']+") # Cleaning for Punctuations and spliting into words 
    words = tokenizer.tokenize(txt)
    
    # Cleaning 
    while True:
        wordsLength = len(words)
        for each in words: 
            if each.isalpha() is False: # Keeping only text
                words.remove(each)
            elif each in stopwords.words('english'): # Cleaning for Stopwords
                words.remove(each)
            elif each in stop_wordsList: # Cleaning for Stopwords from a list
                words.remove(each)
            elif len(each) == 1 and each.upper() not in ['A','I']:
                words.remove(each)
        if len(words)==wordsLength:
            break;
            
                     
    section1_1(words)
    section2(words)
    
    word_count.append(len(words))
    
update()
writeExcel()