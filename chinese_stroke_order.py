import requests
from bs4 import BeautifulSoup
import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

START = 0  # Change this number!
excel_file = r"C:\Users\carls\Downloads\Arts2450.xlsm"
output_folder = r"C:\Users\carls\AppData\Roaming\Anki2\User 1\collection.media"

successful_downloads = 0
unsuccessful_downloads = 0
unsuccessful_gifs = set()
already_downloaded = set()

# Function to download a single GIF
def download_gif(word, output_folder):
    global successful_downloads, unsuccessful_downloads, unsuccessful_gifs
    
    url = f"https://www.strokeorder.com/chinese/{word}"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
    except requests.RequestException:
        print(f"Failed to retrieve the page for word: {word}")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)
        return
    
    soup = BeautifulSoup(response.text, 'html.parser')
    gif_tag = soup.find('img', {'alt': f"{word} Stroke Order Animation"})
    if not gif_tag:
        print(f"Could not find the GIF for word: {word}")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)
        return

    gif_url = gif_tag['src']
    if gif_url.startswith('/'):
        gif_url = 'https://www.strokeorder.com' + gif_url
    
    os.makedirs(output_folder, exist_ok=True)
    file_path = os.path.join(output_folder, f"{word}.gif")

    if os.path.exists(file_path):
        print(f"File already exists for word: {word}, skipping download.")
        return

    try:
        gif_response = requests.get(gif_url, timeout=10)
        gif_response.raise_for_status()
        with open(file_path, 'wb') as f:
            f.write(gif_response.content)
        print(f"Downloaded GIF for word: {word}")
        successful_downloads += 1
    except requests.RequestException:
        print(f"Failed to download GIF for word: {word}")
        unsuccessful_downloads += 1
        unsuccessful_gifs.add(word)

# Function to download GIFs using a ThreadPool
def download_gifs_from_excel(excel_file, sheet_name='Sheet1', output_folder=output_folder):
    global already_downloaded
    already_downloaded = set(os.listdir(output_folder))
    
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    words_column = df.iloc[START:, 0]
    words = []

    for entry in words_column:
        if isinstance(entry, str):
            characters = list(entry)
            words.extend(characters)

    with ThreadPoolExecutor(max_workers=10) as executor:
        executor.map(lambda word: download_gif(word, output_folder), words)

    print(f"Total successful downloads: {successful_downloads}")
    print(f"Total unsuccessful downloads: {unsuccessful_downloads}")
    if unsuccessful_gifs:
        print("Unsuccessful GIFs:", unsuccessful_gifs)

download_gifs_from_excel(excel_file)
