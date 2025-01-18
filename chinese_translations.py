import pandas as pd
from googletrans import Translator
from pypinyin import pinyin, Style
import asyncio

# Function to get Pinyin for sentences with tonal marks
def get_pinyin(text):
    return ' '.join([item[0] for item in pinyin(text, style=Style.TONE)])

# Function to translate sentences
async def translate_text(text):
    translator = Translator()
    translated = await translator.translate(text, src='zh-cn', dest='en')
    return translated.text

# Function to apply async translation to the DataFrame
async def process_translations(df):
    df['Pinyin'] = df['Chinese_Sentence'].apply(get_pinyin)
    # Translate sentences using async
    translations = await asyncio.gather(
        *[translate_text(text) for text in df['Chinese_Sentence']]
    )
    df['English'] = translations
    return df

# Main function to handle asyncio in a modern way
def main():
    # Load your Excel file
    input_file = r"C:\Users\carls\Downloads\Input.xlsx"
    df = pd.read_excel(input_file)

    # Ensure the 'Chinese_Sentence' column exists
    if 'Chinese_Sentence' not in df.columns:
        raise ValueError("Column 'Chinese_Sentence' not found in the input file.")

    # Process translations asynchronously
    df = asyncio.run(process_translations(df))

    # Save the results to a new Excel file
    output_file = 'output_file.xlsx'
    df.to_excel(output_file, index=False)

    print('Processing complete. Results saved to', output_file)

if __name__ == "__main__":
    main()
