# Importing required libraries
from collections import Counter
import pandas as pd
import os

# Define the characters to be dropped
drop_dict = ['，', '\n', '。', '、', '：', '(', ')', '[', ']', '.', ',', ' ', 
             '\u3000', '”', '“', '？', '?', '！', '‘', '’', '…', '「', '」', '；']

# Function to count character occurrences
def count_chars(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
        for char in drop_dict:
            text = text.replace(char, '')
        char_counts = Counter(text)
        return char_counts

# Get user input for file path
file_path = input("Please enter the path to your text file: ")

# Call the function and print the results
char_counts = count_chars(file_path)

# Convert the dictionary to a DataFrame
df = pd.DataFrame(list(char_counts.items()), columns=['Character', 'Count'])

# Sort the DataFrame by 'Count' column in descending order
df = df.sort_values(by='Count', ascending=False)

# Save the DataFrame to an Excel file
output_file = os.path.join(os.path.dirname(file_path), "Word_Count_Single_Charater.xlsx")
df.to_excel(output_file, index=False)

print(f"Results saved to '{output_file}'")
