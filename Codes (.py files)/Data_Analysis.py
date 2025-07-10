import pandas as pd
import numpy as np
import os
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
import re

nltk.download('punkt')

# === PATHS (update if needed) ===
STOPWORDS_FOLDER = "D:/Blackcoffer/StopWords-20250702T160606Z-1-001/StopWords"
MASTERDICTIONARY_FOLDER = "D:/Blackcoffer/MasterDictionary-20250702T160600Z-1-001/MasterDictionary"
ARTICLES_FOLDER = "D:/Blackcoffer/articles"
INPUT_PATH = "D:/Blackcoffer/Input.xlsx"
OUTPUT_STRUCTURE_PATH = "D:/Blackcoffer/Output Data Structure.xlsx"
OUTPUT_PATH = "D:/Blackcoffer/Output.xlsx"

POSITIVE_WORDS_FILE = os.path.join(MASTERDICTIONARY_FOLDER, "positive-words.txt")
NEGATIVE_WORDS_FILE = os.path.join(MASTERDICTIONARY_FOLDER, "negative-words.txt")

# === Load all stopwords from all files in folder ===
stop_words = set()
for fname in os.listdir(STOPWORDS_FOLDER):
    if fname.endswith('.txt'):
        fpath = os.path.join(STOPWORDS_FOLDER, fname)
        with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                word = line.strip().split('|')[0].strip()
                if word:
                    stop_words.add(word.lower())

# === Load positive/negative wordlists, excluding stopwords, ignoring errors that popped up ===
def load_wordlist(filepath, stop_words):
    words = set()
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            w = line.strip().lower()
            if w and w not in stop_words:
                words.add(w)
    return words


positive_words = load_wordlist(POSITIVE_WORDS_FILE, stop_words)
negative_words = load_wordlist(NEGATIVE_WORDS_FILE, stop_words)

# === Syllable counting, complex word, and pronoun helper functions ===
def count_syllables(word):
    word = word.lower()
    if word.endswith(('es', 'ed')) and len(word) > 2:
        word = word[:-2]
    vowels = 'aeiou'
    count = 0
    prev_vowel = False
    for ch in word:
        if ch in vowels:
            if not prev_vowel:
                count += 1
            prev_vowel = True
        else:
            prev_vowel = False
    if word.endswith("e") and not word.endswith(("le", "ue")):
        count -= 1
    if count == 0:
        count = 1
    return count

def is_complex(word):
    return count_syllables(word) > 2

def get_personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    us_as_country = re.findall(r'\bUS\b', text)
    return len(pronouns) - len(us_as_country)

# === Load Output Structure and Input Mapping ===
output_structure = pd.read_excel(OUTPUT_STRUCTURE_PATH)
output_columns = list(output_structure.columns)
input_df = pd.read_excel(INPUT_PATH)
url_lookup = dict(zip(input_df['URL_ID'].astype(str), input_df.to_dict(orient='records')))

# === Process each article ===
results = []

for filename in sorted([f for f in os.listdir(ARTICLES_FOLDER) if f.endswith('.txt')]):
    url_id = filename.replace('.txt', '')
    input_row = url_lookup.get(url_id, {})
    with open(os.path.join(ARTICLES_FOLDER, filename), 'r', encoding='utf-8') as f:
        lines = f.readlines()
        title = lines[0].strip()
        text = " ".join([line.strip() for line in lines[1:]])

    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    words_alpha = [w for w in words if w.isalpha() and w.lower() not in stop_words]
    word_count = len(words_alpha)
    sentence_count = len(sentences)

    # Sentiment
    pos_score = sum(1 for w in words_alpha if w.lower() in positive_words)
    neg_score = sum(1 for w in words_alpha if w.lower() in negative_words)
    polarity_score = (pos_score - neg_score) / ((pos_score + neg_score) + 0.000001)
    subjectivity_score = (pos_score + neg_score) / (word_count + 0.000001)

    # Complex words
    complex_words = [w for w in words_alpha if is_complex(w)]
    complex_word_count = len(complex_words)
    percentage_complex_words = complex_word_count / (word_count + 0.000001)

    # Readability
    avg_sentence_length = word_count / (sentence_count + 0.000001)
    fog_index = 0.4 * (avg_sentence_length + (percentage_complex_words * 100))
    avg_words_per_sentence = avg_sentence_length

    # Syllables
    syllable_per_word = np.mean([count_syllables(w) for w in words_alpha]) if word_count else 0

    # Pronouns
    personal_pronouns = get_personal_pronouns(text)

    # Avg word length
    avg_word_length = np.mean([len(w) for w in words_alpha]) if word_count else 0

    # Output Row (as per output columns order)
    output_row = []
    for col in output_columns:
        # Use values from input_df for input columns, else computed value for output metrics
        col_upper = col.upper().replace(" ", "_")
        if col in input_row:
            output_row.append(input_row[col])
        elif col_upper == "POSITIVE_SCORE":
            output_row.append(pos_score)
        elif col_upper == "NEGATIVE_SCORE":
            output_row.append(neg_score)
        elif col_upper == "POLARITY_SCORE":
            output_row.append(polarity_score)
        elif col_upper == "SUBJECTIVITY_SCORE":
            output_row.append(subjectivity_score)
        elif col_upper == "AVG_SENTENCE_LENGTH":
            output_row.append(avg_sentence_length)
        elif col_upper == "PERCENTAGE_OF_COMPLEX_WORDS":
            output_row.append(percentage_complex_words)
        elif col_upper == "FOG_INDEX":
            output_row.append(fog_index)
        elif col_upper == "AVG_NUMBER_OF_WORDS_PER_SENTENCE":
            output_row.append(avg_words_per_sentence)
        elif col_upper == "COMPLEX_WORD_COUNT":
            output_row.append(complex_word_count)
        elif col_upper == "WORD_COUNT":
            output_row.append(word_count)
        elif col_upper == "SYLLABLE_PER_WORD":
            output_row.append(syllable_per_word)
        elif col_upper == "PERSONAL_PRONOUNS":
            output_row.append(personal_pronouns)
        elif col_upper == "AVG_WORD_LENGTH":
            output_row.append(avg_word_length)
        else:
            output_row.append(None)  # In case of missing/extra columns

    results.append(output_row)

# === Save Output ===
output_df = pd.DataFrame(results, columns=output_columns)
output_df.to_excel(OUTPUT_PATH, index=False)
print("Text analysis complete! Check Output.xlsx for your results.")
