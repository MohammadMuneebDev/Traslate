import openai
import pandas as pd
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# Set up your OpenAI API key
openai.api_key = ''  # Replace with your OpenAI API key

# Dictionary to store translations (cache)
translation_cache = {}

# Function to translate text to English using OpenAI API with caching
def translate_text(text):
    if text in translation_cache:
        return translation_cache[text]
    
    prompt = f"Translate the following Japanese text to English. Do not translate any English text, numbers, or dates: '{text}'"
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=100,
            temperature=0.3,
        )
        translated_text = response.choices[0].message['content'].strip()
        translation_cache[text] = translated_text
        return translated_text
    except Exception as e:
        print(f"Error translating text: {e} for text: '{text}'")
        return text  # Return original text on error

# Function to translate an entire column
def translate_column(column_name, column_data):
    return [translate_text(cell) if isinstance(cell, str) and cell.strip() else cell for cell in column_data]

# Load Excel file
file_path = 'Sawakami Asset Management_Jul 2023 - Jun 2024.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Print original column names
original_columns = df.columns.tolist()
print("Original Columns in DataFrame:", original_columns)

# Translate column headings
translated_columns = [translate_text(col) for col in original_columns]
df.columns = translated_columns

# Print translated column names
print("\nTranslated Columns:")
for col in translated_columns:
    print(col)

# Prompt the user to specify columns to ignore using translated column names
ignore_columns_input = input("\nEnter the translated column names to ignore, separated by commas: ")
ignore_columns = [col.strip() for col in ignore_columns_input.split(',')]  # Split input and remove extra spaces

# Debugging: Print ignore columns
print("\nColumns to ignore:", ignore_columns)

# Prepare for concurrent translation of non-ignored columns
total_columns = sum(1 for col in df.columns if col not in ignore_columns)
with ThreadPoolExecutor() as executor:
    futures = {}
    
    for column in df.columns:
        if column not in ignore_columns:
            futures[executor.submit(translate_column, column, df[column])] = column

    # Wait for all translations to complete and print progress
    completed_columns = 0
    for future in as_completed(futures):
        column_name = futures[future]
        df[column_name] = future.result()
        completed_columns += 1
        print(f"Translated column '{column_name}'. Progress: {completed_columns}/{total_columns} columns completed.")

# Generate output file name by appending "translated" to the base name of the input file
base_name = os.path.basename(file_path)  # Get the base name of the input file
name, ext = os.path.splitext(base_name)  # Split the name and extension
output_file_path = f"{name}_translated{ext}"  # Create new file name

# Save the translated data to a new Excel file
df.to_excel(output_file_path, index=False)
print(f'\nTranslated data has been saved to {output_file_path}')
