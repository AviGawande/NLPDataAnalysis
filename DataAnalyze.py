# import openpyxl  # For reading and writing Excel files
# import nltk
# import pandas as pd
# from nltk.tokenize import word_tokenize, sent_tokenize  # Tokenization
# from collections import Counter  # Counting word occurrences
# from nltk.corpus import cmudict  # Syllable dictionary for syllable count
# import os  # For file operations

# def load_stopwords(filepath):
#     """Loads stopwords from a text file."""
#     stop_words = set()
#     with open(filepath, 'r') as f:
#         stop_words.update([line.strip() for line in f])
#     return stop_words

# def load_positive_negative_words(directory):
#     """Loads positive and negative words from text files."""
#     positive_words = set()
#     negative_words = set()
#     for filename in os.listdir(directory):
#         filepath = os.path.join(directory, filename)
#         with open(filepath, 'r') as f:
#             if filename.startswith("positive"):
#                 positive_words.update([line.strip() for line in f])
#             else:
#                 negative_words.update([line.strip() for line in f])
#     return positive_words, negative_words

# def analyze_text(text, stop_words, positive_words, negative_words):
#     """
#     Analyzes text and computes various metrics.

#     Args:
#       text: The text to analyze.
#       stop_words: Set of stopwords provided.
#       positive_words: Set of positive words.
#       negative_words: Set of negative words.

#     Returns:
#       A dictionary containing the computed metrics.
#     """
#     # Preprocess text
#     text = text.lower()
#     tokens = word_tokenize(text)
#     tokens = [w for w in tokens if w not in stop_words]

#     # Word counts and basic stats
#     word_count = len(tokens)
#     sentence_count = len(sent_tokenize(text))
#     avg_sentence_length = word_count / (sentence_count + 1e-9)  # Avoid division by zero

#     # Complex word handling (assuming 3+ syllables are complex)
#     cmudict_obj = cmudict.dict()
#     complex_words = 0
#     for word in tokens:
#         if any(len(phoneme) >= 3 for phoneme in cmudict_obj.get(word, [])):
#             complex_words += 1

#     # Sentiment analysis (basic approach using positive and negative word counts)
#     positive_score = len([w for w in tokens if w in positive_words])
#     negative_score = len([w for w in tokens if w in negative_words])
#     polarity_score = positive_score - negative_score

#     # Subjectivity score (assuming pronouns indicate subjective content)
#     personal_pronouns = ["i", "me", "my", "mine", "you", "your", "yours", "he", "him", "his", "she", "her", "hers", "it", "its", "we", "us", "our", "ours", "they", "them", "theirs"]
#     subjectivity_score = len([w for w in tokens if w in personal_pronouns]) / word_count

#     # Fog Index (estimating readability)
#     fog_index = 0.4 * ((word_count / sentence_count) + (complex_words / word_count * 100))

#     # Syllable and word length
#     syllables = 0
#     for word in tokens:
#         syllables += sum(1 for phoneme in cmudict_obj.get(word, []) if phoneme[-1] in ['1', '2', '3'])
#     avg_word_length = sum(len(w) for w in tokens) / word_count
#     syllables_per_word = syllables / word_count

#     return {
#         "POSITIVE SCORE": positive_score,
#         "NEGATIVE SCORE": negative_score,
#         "POLARITY SCORE": polarity_score,
#         "SUBJECTIVITY SCORE": subjectivity_score,
#         "AVG SENTENCE LENGTH": avg_sentence_length,
#         "PERCENTAGE OF COMPLEX WORDS": (complex_words / word_count) * 100,
#         "FOG INDEX": fog_index,
#         "AVG NUMBER OF WORDS PER SENTENCE": avg_sentence_length,
#         "COMPLEX WORD COUNT": complex_words,
#         "WORD COUNT": word_count,
#         "SYLLABLE PER WORD": syllables_per_word,
#         "PERSONAL PRONOUNS": len([w for w in tokens if w in personal_pronouns]),
#         "AVG WORD LENGTH": avg_word_length,
#     }

# def perform_data_analysis(input_folder, output_file, stopwords_filepath, master_dir):
#     output_data = []

#     # Load stopwords from file
#     stop_words = load_stopwords(stopwords_filepath)

#     # Load positive and negative words from directory
#     positive_words, negative_words = load_positive_negative_words(master_dir)

#     # Iterate over each text file and compute variables
#     for filename in os.listdir(input_folder):
#         if filename.endswith(".txt"):
#             with open(os.path.join(input_folder, filename), "r", encoding="utf-8") as file:
#                 text = file.read()
#                 metrics = analyze_text(text, stop_words, positive_words, negative_words)
#                 output_data.append({'URL_ID': filename.split('.')[0], **metrics})

#     # Convert the list of dictionaries to a DataFrame
#     df = pd.DataFrame(output_data)

#     # Rearrange the columns to match the specified output structure
#     df = df[['URL_ID', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
#              'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
#              'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
#              'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']]

#     # Save the DataFrame to an Excel file
#     df.to_excel(output_file, index=False)
#     print("Analysis completed and output saved to", output_file)

# def main():
#     input_folder = 'Extracted_Articles'
#     output_file = 'Output Data Structure.xlsx'
#     stopwords_filepath = 'StopWords'
#     master_dir = 'MasterDirectory'

#     perform_data_analysis(input_folder, output_file, stopwords_filepath, master_dir)

# if __name__ == "__main__":
#     main()


import openpyxl  # For reading and writing Excel files
import nltk
import pandas as pd
from nltk.tokenize import word_tokenize, sent_tokenize  # Tokenization
from collections import Counter  # Counting word occurrences
from nltk.corpus import cmudict  # Syllable dictionary for syllable count
import os  # For file operations

def load_stopwords(directory):
    """Loads stopwords from a directory containing multiple text files."""
    stop_words = set()
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r') as f:
            stop_words.update([line.strip() for line in f])
    return stop_words

def load_positive_negative_words(directory):
    """Loads positive and negative words from text files."""
    positive_words = set()
    negative_words = set()
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r') as f:
            if filename.startswith("positive"):
                positive_words.update([line.strip() for line in f])
            else:
                negative_words.update([line.strip() for line in f])
    return positive_words, negative_words

def analyze_text(text, stop_words, positive_words, negative_words):
    """
    Analyzes text and computes various metrics.

    Args:
      text: The text to analyze.
      stop_words: Set of stopwords provided.
      positive_words: Set of positive words.
      negative_words: Set of negative words.

    Returns:
      A dictionary containing the computed metrics.
    """
    # Preprocess text
    text = text.lower()
    tokens = word_tokenize(text)
    tokens = [w for w in tokens if w not in stop_words]

    # Word counts and basic stats
    word_count = len(tokens)
    sentence_count = len(sent_tokenize(text))
    avg_sentence_length = word_count / (sentence_count + 1e-9)  # Avoid division by zero

    # Complex word handling (assuming 3+ syllables are complex)
    cmudict_obj = cmudict.dict()
    complex_words = 0
    for word in tokens:
        if any(len(phoneme) >= 3 for phoneme in cmudict_obj.get(word, [])):
            complex_words += 1

    # Sentiment analysis (basic approach using positive and negative word counts)
    positive_score = len([w for w in tokens if w in positive_words])
    negative_score = len([w for w in tokens if w in negative_words])
    polarity_score = positive_score - negative_score

    # Subjectivity score (assuming pronouns indicate subjective content)
    personal_pronouns = ["i", "me", "my", "mine", "you", "your", "yours", "he", "him", "his", "she", "her", "hers", "it", "its", "we", "us", "our", "ours", "they", "them", "theirs"]
    subjectivity_score = len([w for w in tokens if w in personal_pronouns]) / word_count

    # Fog Index (estimating readability)
    fog_index = 0.4 * ((word_count / sentence_count) + (complex_words / word_count * 100))

    # Syllable and word length
    syllables = 0
    for word in tokens:
        syllables += sum(1 for phoneme in cmudict_obj.get(word, []) if phoneme[-1] in ['1', '2', '3'])
    avg_word_length = sum(len(w) for w in tokens) / word_count
    syllables_per_word = syllables / word_count

    return {
        "POSITIVE SCORE": positive_score,
        "NEGATIVE SCORE": negative_score,
        "POLARITY SCORE": polarity_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_length,
        "PERCENTAGE OF COMPLEX WORDS": (complex_words / word_count) * 100,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_sentence_length,
        "COMPLEX WORD COUNT": complex_words,
        "WORD COUNT": word_count,
        "SYLLABLE PER WORD": syllables_per_word,
        "PERSONAL PRONOUNS": len([w for w in tokens if w in personal_pronouns]),
        "AVG WORD LENGTH": avg_word_length,
    }

def perform_data_analysis(input_folder, output_file, stopwords_directory):
    output_data = []

    # Load stopwords from directory
    stop_words = load_stopwords(stopwords_directory)

    # Load positive and negative words from directory
    positive_words, negative_words = load_positive_negative_words(stopwords_directory)

    # Iterate over each text file and compute variables
    for filename in os.listdir(input_folder):
        if filename.endswith(".txt"):
            with open(os.path.join(input_folder, filename), "r", encoding="utf-8") as file:
                text = file.read()
                metrics = analyze_text(text, stop_words, positive_words, negative_words)
                output_data.append({'URL_ID': filename.split('.')[0], **metrics})

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(output_data)

    # Rearrange the columns to match the specified output structure
    df = df[['URL_ID', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
             'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
             'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
             'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']]

    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False)
    print("Analysis completed and output saved to", output_file)

def main():
    input_folder = 'Extracted_Articles'
    output_file = 'Output Data Structure.xlsx'
    stopwords_directory = 'StopWords'

    perform_data_analysis(input_folder, output_file, stopwords_directory)

if __name__ == "__main__":
    main()
