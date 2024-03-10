from docx import Document
import csv
import os
'''
requirements:

1. put your word docs you want to process in word_docs_to_process file
2. rename combined_tables.csv column names to your desired column names

this script only works if you have the same table in multiple word docs 
with the same table schema and it's in the same position in each document
'''

## User input
# Set the position of the table in the word doc you want to extract from all the words docs(0 means the first table ect)
table_position = 2

# Directory containing the Word documents
word_docs_folder = "word_docs_to_process"

# Flag to track if it's the first document being processed
first_doc_processed = False

for filename in os.listdir(word_docs_folder):
    if filename.endswith(".docx"):
        # Set the name of the Word doc
        word_doc = os.path.join(word_docs_folder, filename)
        word_doc_name = filename.split('.')[0]

        # Open the Word document
        doc = Document(word_doc)
        # Access the first table in the document
        table = doc.tables[table_position]
        # Create a single CSV file to combine all tables
        combined_csv_filename = f"output/combined_tables.csv"
        with open(combined_csv_filename, 'a', newline='', encoding='utf-8') as combined_csvfile:
            writer = csv.writer(combined_csvfile)
                
            # Flag to skip the header after the first document
            skip_header = False
            
            # Iterate through rows in the table
            for row in table.rows:
                if not skip_header:
                    skip_header = True
                    continue  # Skip the header row after the first document
                
                # Extract data from each cell in the row
                row_data = [cell.text.strip() for cell in row.cells]
                # Append the filename as an additional item in the row
                row_data.append(word_doc_name)
                # Write the row to the CSV file
                writer.writerow(row_data)
            
        print(f"Table from position {table_position} from {word_doc_name}, appended to '{combined_csv_filename}'")