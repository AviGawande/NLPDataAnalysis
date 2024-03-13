import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

def extract_article_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        article_title = soup.find('title').get_text()
        article_text = ''
        for paragraph in soup.find_all('p'):
            article_text += paragraph.get_text() + '\n'
        return article_title, article_text
    except Exception as e:
        print(f"Error extracting article from {url}: {e}")
        return None, None

def main():
    input_file = "Input.xlsx"
    output_folder = "Extracted_Articles"

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        df = pd.read_excel(input_file, engine='openpyxl')  # Specify engine='openpyxl'
        for index, row in df.iterrows():
            url = row['URL']
            url_id = row['URL_ID']
            article_title, article_text = extract_article_text(url)
            if article_title and article_text:
                with open(os.path.join(output_folder, f"{url_id}.txt"), 'w', encoding='utf-8') as f:
                    f.write(f"{article_title}\n\n{article_text}")
                print(f"Article extracted and saved for {url_id}")
    except Exception as e:
        print(f"Error processing {input_file}: {e}")

if __name__ == "__main__":
    main()


