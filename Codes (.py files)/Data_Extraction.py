import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
from tqdm import tqdm

# Step 1: Read URLs from Input.xlsx
input_df = pd.read_excel('Input.xlsx')

# Step 2: Prepare output directory
os.makedirs('articles')

# Step 3: Prepare error log file (with error type column)
error_log_path = 'errors.txt'
with open(error_log_path, 'w', encoding='utf-8') as ef:
    ef.write('URL_ID\tURL\tError_Type\tError_Message\n')  # Header row

# Step 4: Function to extract title and main article text
def extract_title_and_text(url):
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()  # Raises HTTPError for bad responses
        soup = BeautifulSoup(response.content, 'html.parser')

        # Get title
        if soup.title and soup.title.string:
            title = soup.title.string.strip()
        elif soup.find('h1'):
            title = soup.find('h1').get_text().strip()
        else:
            title = ''

        # Get article text
        article_tag = soup.find('article')
        if article_tag:
            paragraphs = article_tag.find_all('p')
            article_text = '\n'.join([p.get_text() for p in paragraphs])
        else:
            paragraphs = soup.find_all('p')
            article_text = '\n'.join([p.get_text() for p in paragraphs])

        article_text = article_text.strip()
        return title, article_text

    except requests.exceptions.Timeout as e:
        raise Exception("TimeoutError: " + str(e))
    except requests.exceptions.HTTPError as e1:
        raise Exception("HTTPError: " + str(e1))
    except Exception as e2:
        raise Exception("ParsingOrOtherError: " + str(e2))

# Step 5: Iterate and save each article to text file, or log errors with type
for i, row in tqdm(input_df.iterrows(), total=len(input_df)):
    url_id = str(row['URL_ID'])
    url = row['URL']
    try:
        title, text = extract_title_and_text(url)

        # Clean up: strip spaces from each line
        clean_title = title.strip()
        clean_lines = [line.strip() for line in text.split('\n')]
        clean_text = '\n'.join([line for line in clean_lines if line])  # Remove empty lines

        filename = "articles/" + url_id + ".txt"
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(clean_title + '\n' + clean_text)
    except Exception as e:
        error_message = str(e)
        # Determine error type (basic approach)
        if "TimeoutError" in error_message:
            error_type = "Timeout"
        elif "HTTPError" in error_message:
            error_type = "HTTP"
        elif "ParsingOrOtherError" in error_message:
            error_type = "ParsingOrOther"
        else:
            error_type = "Unknown"
        # Traditional formatting, with all error columns
        log_line = url_id + "\t" + url + "\t" + error_type + "\t" + error_message + "\n"
        with open(error_log_path, 'a', encoding='utf-8') as ef:
            ef.write(log_line)

print("Extraction complete! Check the 'articles' directory for your files, and errors.txt for a summary of any issues.")
