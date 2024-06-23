import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import syllapy
import re
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
nltk.download('punkt')


# accessing stop words
folder_path = "C:/Users/ASUS/Desktop/Web Scraping/StopWords"
total_stop_words = ""

for stop_files in os.listdir(folder_path):
    with open(os.path.join(folder_path, stop_files), 'r') as f:
        file_content = f.read()
        total_stop_words += file_content
        

split_in_lines = total_stop_words.split('\n')

cleaned_stop_words = []

for line in split_in_lines:
    if "|" in line:
        stop_word = line.split(" | ")[0]
        cleaned_stop_words.append(stop_word)
    else:
        cleaned_stop_words.append(line)

# accessing positive and negative words
with open("C:/Users/ASUS/Desktop/Web Scraping/MasterDictionary/positive-words.txt", "r") as f:
    positive_words = f.read()

with open("C:/Users/ASUS/Desktop/Web Scraping/MasterDictionary/negative-words.txt", "r") as f:
    negative_words = f.read()


# function to save a web page in the form of html file
def convert_to_html(url, path):
    r = requests.get(url)
    with open(path, "w", encoding='utf-8') as f:
        f.write(r.text)

# function to extract only title and content of the article
def convert_to_text(html_folder, text_folder, url_id):
    with open(os.path.join(html_folder,f"{url_id}.html"), "r", encoding='utf-8') as f:
        file = f.read()

    soup = BeautifulSoup(file, 'html.parser')

    title = soup.h1
    if title is not None:
        title_text = title.get_text()
    else:
        title_text = ""

    outer = soup.find(class_ = "td-post-content")
    remove = soup.find(class_ = "wp-block-preformatted")

    content = ""
    if outer is not None and remove is not None:
        for child in outer.children:
            if child != remove:
                content += child.text + '\n'

    final = title_text + "\n" + content

    with open (os.path.join(text_folder,f"{url_id}.txt"),'w', encoding='utf-8') as f:
        f.write(final)

# function to get POSITIVE SCORE
def positive(cleaned_words):
    score = 0

    for word in cleaned_words:
        if word in positive_words:
            score = score + 1
            
    return score

# function to get NEGATIVE SCORE
def negative(cleaned_words):
    score = 0

    for word in cleaned_words:
        if word in negative_words:
            score = score + 1
            
    return score

# function to get POLARITY SCORE
def polarity(positive_score,negative_score):
    score = (positive_score - negative_score)/((positive_score + negative_score) + 0.000001)
    return score

# function to get SUBJECTIVITY SCORE
def subjectivity(positive_score,negative_score,cleaned_words):
    score = (positive_score + negative_score)/((len(cleaned_words)) + 0.000001)
    return score

# function to get COMPLEX WORD COUNT
def complex_word(words):
    count= 0

    for word in words:
        if syllapy.count(word) > 2:
            count = count + 1
    
    return count

# function to get SYLLABLE PER WORD
def syllable(words):
    total_syllables = 0

    for word in words:
        total_syllables = total_syllables + syllapy.count(word)

    return total_syllables/len(words)

# function to get COUNT OF PERSONAL PRONOUNS
def pronouns(text):
    x = '\\b(I|me|my|mine|myself|you|your|yours|yourself|he|him|his|himself|she|her|hers|herself|it|its|itself|we|us|our|ours|ourselves|yourselves|they|them|their|theirs|themselves)\\b'
    matches = re.findall(x, text)

    return len(matches)

# function to get AVG WORD LENGTH
def word_length(words):
    total_length = 0

    for word in words:
        total_length = total_length + len(word)

    return total_length/len(words)

# function to use all functions for text analysis
def text_analysis(text):
      
    words = word_tokenize(text)
    sentences = sent_tokenize(text)

    cleaned_words = []

    for word in words:
        if word not in cleaned_stop_words:
            cleaned_words.append(word)

    if len(words) != 0 and len(sentences) != 0:    
        positive_score = positive(cleaned_words)
        negative_score = negative(cleaned_words)
        polarity_score = polarity(positive_score,negative_score)
        subjectivity_score = subjectivity(positive_score,negative_score,cleaned_words)
        avg_sentence_length = len(words)/len(sentences)
        complex_word_count = complex_word(words)
        percentage_of_complex_words = complex_word_count/len(words)
        fog_index = 0.4*(avg_sentence_length + percentage_of_complex_words)
        avg_number_of_words_per_sentence = len(words)/len(sentences)
        word_count = len(cleaned_words)
        syllable_per_word = syllable(words)
        personal_pronouns = pronouns(file_text)
        avg_word_length = word_length(words)
        
    else:
        positive_score = 0
        negative_score = 0
        polarity_score = 0
        subjectivity_score = 0
        avg_sentence_length = 0
        complex_word_count = 0
        percentage_of_complex_words = 0
        fog_index = 0
        avg_number_of_words_per_sentence = 0
        word_count = 0
        syllable_per_word = 0
        personal_pronouns = 0
        avg_word_length = 0
        
    return [positive_score,negative_score,polarity_score,subjectivity_score,avg_sentence_length,complex_word_count,percentage_of_complex_words,fog_index,avg_number_of_words_per_sentence,word_count,syllable_per_word,personal_pronouns,avg_word_length]


# reading input.xlsx
file = pd.read_excel("C:/Users/ASUS/Desktop/Web Scraping/Input.xlsx")   
df = pd.DataFrame()
 
# extracting URL_ID and URL from each row of excel file to perform text analysis
for i in range (0, len(file)):
    url_id = file['URL_ID'][i]
    url = file['URL'][i]
    convert_to_html(url, f"C:/Users/ASUS/Desktop/Web Scraping/HTML Files/{url_id}.html")
    convert_to_text("C:/Users/ASUS/Desktop/Web Scraping/HTML Files", "C:/Users/ASUS/Desktop/Web Scraping/Text Files", url_id)
    
    with open(f"C:/Users/ASUS/Desktop/Web Scraping/Text Files/{url_id}.txt", "r", encoding='utf-8') as f:
        file_text = f.read()
    
    # calling text_analysis function to perform text analysis on contents of each text file
    output = text_analysis(file_text)
    
    # making a dataframe of the desired format
    output.insert(0,url_id)
    output.insert(1,url)
    
    columns = ['URL_ID','URL','POSITIVE SCORE','NEGATIVE SCORE','POLARITY SCORE','SUBJECTIVITY SCORE','AVG SENTENCE LENGTH','PERCENTAGE OF COMPLEX WORDS','FOG INDEX','AVG NUMBER OF WORDS PER SENTENCE','COMPLEX WORD COUNT','WORD COUNT','SYLLABLE PER WORD','PERSONAL PRONOUNS','AVG WORD LENGTH']
    
    data = pd.DataFrame([output], columns = columns)
    df = pd.concat([df,data], ignore_index = True)

# exporting dataframe to a excel file
df.to_excel("C:/Users/ASUS/Desktop/Web Scraping/Output Data Structure.xlsx", index = False)