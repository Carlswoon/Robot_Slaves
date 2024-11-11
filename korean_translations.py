import pandas as pd
from googletrans import Translator
from korean_romanizer import Romanizer

# Function to romanize Korean sentences
def romanize_korean(text):
    romanizer = Romanizer(text)
    return romanizer.romanize()

# Function to translate sentences
def translate_text(text, src_lang, dest_lang='en'):
    translator = Translator()
    translated = translator.translate(text, src=src_lang, dest=dest_lang)
    return translated.text

# Load your Excel file
input_file = r'C:\Users\carls\Downloads\testing..xlsx'  # Use raw string literal or forward slashes
df = pd.read_excel(input_file)

# Ensure the 'Korean_Sentence' column exists
if 'Korean_Sentence' not in df.columns:
    raise ValueError("Column 'Korean_Sentence' not found in the input file.")

# Process each sentence
df['Romanized'] = df['Korean_Sentence'].apply(romanize_korean)
df['English'] = df['Korean_Sentence'].apply(lambda x: translate_text(x, src_lang='ko'))

# Save the results to a new Excel file
output_file = 'output_file.xlsx'  # Adjust as needed
df.to_excel(output_file, index=False)

print('Processing complete. Results saved to', output_file)

