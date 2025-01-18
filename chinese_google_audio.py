import requests
import os
import pandas as pd
import urllib.parse
import re
from concurrent.futures import ThreadPoolExecutor

START = 0  # Change this number!
excel_file = r"C:\Users\carls\Downloads\Arts2450.xlsm"
output_folder = r"C:\Users\carls\AppData\Roaming\Anki2\User 1\collection.media"

# Utility to sanitize filenames
def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '？', filename)

# Utility to clean text for URL
def clean_text_for_url(text):
    return re.sub(r'[()]', '', text).strip()

# Function to download a single audio file
def download_audio(text, output_folder, processed_texts):
    cleaned_text_for_url = clean_text_for_url(text)
    
    # Check if already processed
    if cleaned_text_for_url in processed_texts:
        return (False, text, "Duplicate entry")
    
    processed_texts.add(cleaned_text_for_url)
    
    # Construct URL
    base_url = "https://translate.google.com/translate_tts"
    params = {
        "ie": "UTF-8",
        "client": "tw-ob",
        "tl": "zh-CN",  # Target language for Chinese
        "q": cleaned_text_for_url
    }
    url = f"{base_url}?{urllib.parse.urlencode(params)}"
    
    # Prepare file path
    os.makedirs(output_folder, exist_ok=True)
    safe_text = sanitize_filename(text)
    file_path = os.path.join(output_folder, f"{safe_text}.mp3")
    
    if os.path.exists(file_path):
        return (False, text, "File already exists")
    
    # Download the audio
    try:
        response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        response.raise_for_status()
        with open(file_path, 'wb') as f:
            f.write(response.content)
        return (True, text, "Success")
    except requests.RequestException as e:
        return (False, text, f"Failed with error: {str(e)}")

# Function to download audios concurrently
def download_audios_from_excel(excel_file, sheet_name='Sheet1', output_folder=output_folder):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    text_column = df.iloc[START:, 0]
    
    processed_texts = set()
    results = []

    # Use ThreadPoolExecutor for concurrent downloads
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {
            executor.submit(download_audio, str(entry), output_folder, processed_texts): entry
            for entry in text_column if isinstance(entry, str)
        }
        
        # Collect results
        for future in futures:
            success, text, message = future.result()
            results.append((success, text, message))
    
    # Summarize results
    success_count = sum(1 for r in results if r[0])
    fail_count = len(results) - success_count
    failed_texts = [r[1] for r in results if not r[0]]

    print(f"Total successful downloads: {success_count}")
    print(f"Total failed downloads: {fail_count}")
    if failed_texts:
        print("Failed to download audio for the following texts:")
        for failed_text in failed_texts:
            print(f"- {failed_text}")

download_audios_from_excel(excel_file, output_folder=output_folder)
