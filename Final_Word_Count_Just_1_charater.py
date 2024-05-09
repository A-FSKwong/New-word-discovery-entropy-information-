# Importing required libraries
from collections import Counter
import pandas as pd
import os

# Define the characters to be dropped
drop_dict = [u'，', u'\n', u'。', u'、', u'：', u'(', u')', u'[', u']', u'.', u',', u' ', 
             u'1', u'2', u'3', u'4', u'5', u'6', u'7', u'8', u'9', u'0',
             u'-', u'—', u'"', u'*', u'/', u':', u'@', u'～', u'《', u'》',
             u'『', u'』', u'+', u'○', u'◎', u'０', u'〇', u'２', u'4', u'５',
             u'６', u'８', u'9', u'a', u'A', u'B', u'b', u'c', u'C', u'd',
             u'e', u'E', u'f', u'F', u'g', u'h', u'H', u'i', u'j', u'k',
             u'l', u'L', u'm', u'M', u'n', u'N', u'o', u'p', u'P', u'Q',
             u'q', u'r', u'R', u's', u'S', u't', u'T', u'u', u'U', u'v',
             u'V', u'w', u'x', u'y', u'Y', u'Z', u'G', u'D', u'I', u'J',
             u'K', u'O', u'X', u';', u'&', u'%', u'$', u'、', u'。', u'，',u'（', u'）', u'；', u'４', 
             u'\u3000', u'”', u'“', u'？', u'?', u'！', u'‘', u'’', u'…', u'「', u'」']

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
