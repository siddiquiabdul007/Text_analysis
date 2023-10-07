import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.probability import FreqDist
from nltk.corpus import stopwords
import string
import pandas as pd
import re

# Download necessary NLTK data (you only need to run this once)
nltk.download('punkt')
nltk.download('stopwords')

# Function to perform text analysis
def perform_text_analysis(content):
    # Tokenize the text (split into words)
    words = word_tokenize(content.lower())

    # Remove punctuation and stopwords
    table = str.maketrans('', '', string.punctuation)
    words = [word.translate(table) for word in words]
    stop_words = set(stopwords.words('english'))
    words = [word for word in words if word.isalpha() and word not in stop_words]

    # Perform frequency distribution analysis
    freq_dist = FreqDist(words)

    # Get the most common words
    most_common_words = [word for word, freq in freq_dist.most_common(10)]

    # Perform text analysis for POSITIVE SCORE, NEGATIVE SCORE, POLARITY SCORE, and SUBJECTIVITY SCORE
    positive_words = set(["good", "positive", "excellent"])  # Add positive words here
    negative_words = set(["bad", "negative", "poor"])  # Add negative words here

    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)

    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    polarity_score = max(-1, min(polarity_score, 1))  # Limit the range to -1 to +1

    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)

    # Analysis of Readability
    total_words = len(words)
    total_sentences = len(sent_tokenize(content))
    avg_sentence_length = total_words / total_sentences

    complex_words = [word for word in words if len(word) > 2]
    percentage_complex_words = len(complex_words) / total_words
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    # Other variables
    complex_word_count = len(complex_words)
    word_count = total_words

    # Syllable Count Per Word
    vowels = "aeiouy"
    syllable_per_word = sum([len(re.findall(r'[aeiouy]+', word)) for word in words])

    # Personal Pronouns
    personal_pronouns = len(re.findall(r'\b(?:I|we|my|ours|us)\b', content, re.IGNORECASE))

    # Average Word Length
    total_characters = sum(len(word) for word in words)
    avg_word_length = total_characters / total_words

    return {
        "POSITIVE_SCORE": positive_score,
        "NEGATIVE_SCORE": negative_score,
        "POLARITY_SCORE": polarity_score,
        "SUBJECTIVITY_SCORE": subjectivity_score,
        "AVG_SENTENCE_LENGTH": avg_sentence_length,
        "PERCENTAGE_OF_COMPLEX_WORDS": percentage_complex_words,
        "FOG_INDEX": fog_index,
        "AVG_NUMBER_OF_WORDS_PER_SENTENCE": total_words / total_sentences,
        "COMPLEX_WORD_COUNT": complex_word_count,
        "WORD_COUNT": word_count,
        "SYLLABLE_PER_WORD": syllable_per_word,
        "PERSONAL_PRONOUNS": personal_pronouns,
        "AVG_WORD_LENGTH": avg_word_length,
        "MOST_COMMON_WORDS": most_common_words,
    }

# Read the list of URLs and URL_IDs from the Input.xlsx file
input_file = "Input.xlsx"
df = pd.read_excel(input_file)

# Create an empty list to store the results
output_data = []

# Loop through each URL_ID and perform text analysis
for index, row in df.iterrows():
    url_id = row["URL_ID"]
    url = row["URL"]

    try:
        # Read the content from the text file with the URL_ID as the filename
        filename = f"{url_id}.txt"
        with open(filename, "r", encoding="utf-8") as file:
            content = file.read()

        # Perform text analysis
        text_analysis_results = perform_text_analysis(content)

        # Append the results to the output_data list
        output_data.append({"URL_ID": url_id, **text_analysis_results})

    except FileNotFoundError:
        print(f"Data file not found for URL_ID: {url_id} - Skipping.")

# Create a DataFrame from the output_data list
output_df = pd.DataFrame(output_data)

# Save the results to the output Excel file
output_file = "Output Data Structure.xlsx"
output_df.to_excel(output_file, index=False)

print(f"Text analysis results saved to {output_file}.")
