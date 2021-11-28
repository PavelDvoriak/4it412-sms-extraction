import os
import sys
import xls_output as xo

dir = os.fsencode('./theses/')

all_theses_data = dict()
TERMS = ('softwarové řešení', )

def main():
    global all_theses_data

    for file in os.listdir(dir):
        filename = os.fsdecode(file)
        theses_data = dict()
        lines_count = 0

        with open(f'./theses/{filename}', encoding='utf-8') as f:

            # Loop through each line of the file
            for line in f:
                # Remove the leading spaces and newline character
                if line == '\n':
                    continue
                
                line = line.strip()
            
                lines_count += 1

                # Convert the characters in line to 
                # lowercase to avoid case mismatch
                line = line.lower()

                # Check if line contains terms
                for term in TERMS:
                    if term in line:
                        line = line.replace(term, '')
                        add_to_dict(theses_data, term)

                # Split the line into words
                words = line.split(" ")
            
                # Iterate over each word in line
                for word in words:
                    if any(map(str.isdigit, word)):
                        continue
                    
                    # Check if the word is already in dictionary
                    add_to_dict(theses_data, word)

                theses_data['chapters_count'] = lines_count

            all_theses_data[filename[:-4]] = theses_data
            total_data = dict()

    for file_key in all_theses_data:
        for word_key in all_theses_data[file_key]:
            word_count = all_theses_data[file_key][word_key]
            if word_key in total_data:
                total_data[word_key] = total_data[word_key] + word_count
            else:
                total_data[word_key] = word_count
    all_theses_data['total'] = total_data
    xo.output_to_xls(all_theses_data)

def add_to_dict(dictionary: dict, word: str) -> None:
    if word in dictionary:
        # Increment count of word by 1
        dictionary[word] = dictionary[word] + 1
    else:
        # Add the word to dictionary with count 1
        dictionary[word] = 1


main()
