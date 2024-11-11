import requests
import os
import time
import pandas as pd
import urllib.parse
import re

START = 1  # Change this number!
# Specify the path to your Excel file
excel_file = r"C:\Users\carls\Downloads\Chinese word vocab populate.xlsx"
output_folder = r"C:\Users\carls\AppData\Roaming\Anki2\Carlson\collection.media"

def sanitize_filename(filename):
    # Replace characters that are not allowed in filenames, preserving brackets
    return re.sub(r'[<>:"/\\|?*]', '？', filename)

def clean_text_for_url(text):
    # Remove brackets but keep the content inside for URL
    cleaned_text = re.sub(r'[()]', '', text).strip()
    return cleaned_text

def download_audio(text, output_folder='audio_files', processed_texts=None):
    # Clean the text for URL before using it
    cleaned_text_for_url = clean_text_for_url(text)
    
    # Check if the text has already been processed
    if cleaned_text_for_url in processed_texts:
        print(f"Duplicate entry found, counting as failure for text: {text}")
        return False  # Indicate a failure due to duplicate

    # Add the cleaned text to the set of processed texts
    processed_texts.add(cleaned_text_for_url)
    
    # Construct the URL for the Google Translate TTS service
    base_url = "https://translate.google.com/translate_tts"
    params = {
        "ie": "UTF-8",
        "client": "tw-ob",
        "tl": "zh-CN",  # Target language for Korean
        "q": cleaned_text_for_url
    }
    url = f"{base_url}?{urllib.parse.urlencode(params)}"
    
    # Use a filename safe for characters, preserving brackets in filenames
    os.makedirs(output_folder, exist_ok=True)
    safe_text = sanitize_filename(text)
    file_path = os.path.join(output_folder, f"{safe_text}.mp3")

    # Check if the file already exists
    if os.path.exists(file_path):
        print(f"File already exists for text: {text}, skipping download.")
        return False  # Indicate a failure due to existing file

    # Send a request to get the audio
    response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"})
    if response.status_code == 200:
        with open(file_path, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded audio for text: {text}")
        return True
    else:
        print(f"Failed to download audio for text: {text}")
        return False

def download_audios_from_excel(excel_file, sheet_name='Sheet1', output_folder=r"C:\Users\carls\AppData\Roaming\Anki2\Carlson\collection.media"):
    # Read the Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    
    # Assuming text is in the first column
    text_column = df.iloc[START:, 0]
    
    success_count = 0
    fail_count = 0
    failed_texts = []  # List to keep track of failed downloads
    processed_texts = set()  # Set to track processed texts (including duplicates)
    
    for entry in text_column:
        # Process the entire cell content as a single string
        if isinstance(entry, str):
            success = download_audio(entry, output_folder, processed_texts)
            if success:
                success_count += 1
            else:
                fail_count += 1
                failed_texts.append(entry)  # Add the failed text to the list
            time.sleep(1)  # Sleep to avoid making too many requests in a short time

    # Print out the final counts
    print(f"Total successful downloads: {success_count}")
    print(f"Total failed downloads: {fail_count}")
    
    # Print out the list of failed texts if there are any
    if fail_count > 0:
        print("Failed to download audio for the following texts:")
        for failed_text in failed_texts:
            print(f"- {failed_text}")

download_audios_from_excel(excel_file, output_folder=output_folder)
