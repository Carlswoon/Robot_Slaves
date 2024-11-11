import requests
from bs4 import BeautifulSoup
import os
import time
import pandas as pd

START = 1  # Change this number!
# Specify the path to your Excel file
excel_file = r"C:\Users\carls\Downloads\Chinese word vocab populate.xlsx"
# Specify the output folder
output_folder = r"C:\Users\carls\AppData\Roaming\Anki2\Carlson\collection.media"

successful_downloads = 0
unsuccessful_downloads = 0
unsuccessful_gifs = set()

def download_gif(word, output_folder='gifs'):
    global successful_downloads, unsuccessful_downloads, unsuccessful_gifs
    
    # Construct the URL for the page containing the GIF
    url = f"https://www.strokeorder.com/chinese/{word}"
    
    # Get the HTML content of the page
    response = requests.get(url)
    if response.status_code != 200:
        print(f"Failed to retrieve the page for word: {word}")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)
        return
    
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Find the GIF URL in the page
    gif_tag = soup.find('img', {'alt': f"{word} Stroke Order Animation"})
    if not gif_tag:
        print(f"Could not find the GIF for word: {word}")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)
        return

    gif_url = gif_tag['src']
    
    # Construct the full GIF URL
    if gif_url.startswith('/'):
        gif_url = 'https://www.strokeorder.com' + gif_url
    
    # Set the output file path
    os.makedirs(output_folder, exist_ok=True)
    file_path = os.path.join(output_folder, f"{word}.gif")

    # Check if the file already exists
    if os.path.exists(file_path):
        print(f"File already exists for word: {word}, skipping download.")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)
        return

    # Download the GIF
    gif_response = requests.get(gif_url)
    if gif_response.status_code == 200:
        with open(file_path, 'wb') as f:
            f.write(gif_response.content)
        print(f"Downloaded GIF for word: {word}")
        successful_downloads += 1
    else:
        print(f"Failed to download GIF for word: {word}")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)

def download_gifs_from_excel(excel_file, sheet_name='Sheet1', output_folder=r"C:\Users\carls\AppData\Roaming\Anki2\Carlson\collection.media"):
    # Read the Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    
    # Assuming words are in the first column
    words_column = df.iloc[START:, 0]
    
    for entry in words_column:
        # Convert the entry to a string and split into individual characters
        if isinstance(entry, str):
            characters = list(entry)  # Split the string into individual characters
            for char in characters:
                download_gif(char, output_folder)
                time.sleep(1)  # Sleep to avoid making too many requests in a short time

    # Print summary of download results
    print(f"Total successful downloads: {successful_downloads}")
    print(f"Total unsuccessful downloads: {unsuccessful_downloads}")
    if unsuccessful_gifs:
        print("Unsuccessful GIFs:")
        for gif in unsuccessful_gifs:
            print(gif)

download_gifs_from_excel(excel_file, output_folder=output_folder)
