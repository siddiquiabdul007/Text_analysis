import requests
from bs4 import BeautifulSoup
import pandas as pd

# Read the list of URLs and URL_IDs from the Excel file
input_file = "input.xlsx"
df = pd.read_excel(input_file)

# Loop through each URL and extract the article heading and text
for index, row in df.iterrows():
    url = row["URL"]
    url_id = row["URL_ID"]

    try:
        # Send a GET request to the URL and fetch the HTML content
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for non-200 status codes
        html_content = response.content

        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')

        # Extract the article heading
        article_heading = soup.find('h1').text.strip()

        # Extract the article text
        article_text_elements = soup.find_all('p')
        article_text = "\n".join([element.text.strip() for element in article_text_elements])

        # Save the extracted data as a text file with the URL_ID as the filename
        filename = f"{url_id}.txt"
        with open(filename, "w", encoding="utf-8") as file:
            file.write(f"Article Heading for URL: {url}\n")
            file.write(article_heading + "\n\n")
            file.write(f"Article Text for URL: {url}\n")
            file.write(article_text + "\n")

        print(f"Data extracted for URL_ID: {url_id} and saved as {filename}")
    except Exception as e:
        print(f"Sorry, data extraction failed for URL_ID: {url_id} - Relevant data unavailable.")
        continue

    print("----------------------------------------------------")
