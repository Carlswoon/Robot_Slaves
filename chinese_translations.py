import pandas as pd
from googletrans import Translator
from pypinyin import pinyin, Style

# Function to get Pinyin for sentences with tonal marks
def get_pinyin(text):
    return ' '.join([item[0] for item in pinyin(text, style=Style.TONE)])

# Function to translate sentences
def translate_text(text):
    translator = Translator()
    translated = translator.translate(text, src='zh-cn', dest='en')
    return translated.text

# Load your Excel file
input_file = r'C:\Users\carls\Downloads\Chinese word vocab populate.xlsx'  # Use raw string literal or forward slashes
df = pd.read_excel(input_file)

# Ensure the 'Chinese_Sentence' column exists
if 'Chinese_Sentence' not in df.columns:
    raise ValueError("Column 'Chinese_Sentence' not found in the input file.")

# Process each sentence
df['Pinyin'] = df['Chinese_Sentence'].apply(get_pinyin)
df['English'] = df['Chinese_Sentence'].apply(translate_text)

# Save the results to a new Excel file
output_file = 'output_file.xlsx'  # Adjust as needed
df.to_excel(output_file, index=False)

print('Processing complete. Results saved to', output_file)
