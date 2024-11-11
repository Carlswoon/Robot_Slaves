import pandas as pd
from pydub import AudioSegment, silence
import speech_recognition as sr
import os

def load_excel(file_path):
    return pd.read_excel(file_path)

def split_audio(file_path, silence_thresh=-40, min_silence_len=500):
    audio = AudioSegment.from_mp3(file_path)
    audio_chunks = silence.split_on_silence(audio, min_silence_len=min_silence_len, silence_thresh=silence_thresh)
    return audio_chunks

def recognize_speech(audio_chunk):
    recognizer = sr.Recognizer()
    with sr.AudioFile(audio_chunk) as source:
        audio = recognizer.record(source)
        try:
            text = recognizer.recognize_google(audio, language='ko-KR')
        except sr.UnknownValueError:
            text = None
        except sr.RequestError:
            text = None
    return text

def main(excel_file, audio_file, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Load Excel data
    df = load_excel(excel_file)
    korean_sentences = df['Korean Sentences'].tolist()  # Adjust column name if needed

    # Split audio file
    audio_chunks = split_audio(audio_file)

    # Process each audio chunk
    for i, chunk in enumerate(audio_chunks):
        chunk_file = f"temp_chunk_{i}.wav"
        chunk.export(chunk_file, format="wav")
        
        recognized_text = recognize_speech(chunk_file)
        if recognized_text:
            # Find the corresponding sentence in the Excel data
            for sentence in korean_sentences:
                if recognized_text in sentence:
                    output_file = os.path.join(output_dir, f"{sentence}.mp3")
                    chunk.export(output_file, format="mp3")
                    print(f"Exported: {output_file}")
                    break
        
        os.remove(chunk_file)  # Clean up temporary files

if __name__ == "__main__":
    excel_file = r"C:\Users\carls\Downloads\testing..xlsx"
    audio_file = r"C:\Users\carls\Downloads\Korean\Conversational Korean\page 017 First Meetings.mp3" 
    output_dir = r"C:\Users\carls\output_audio"
    main(excel_file, audio_file, output_dir)