from docx import Document
import csv

## User input
# Set the name of the word doc
word_doc = 'example_doc.docx'
word_doc_name = word_doc.split('.')[0]

# Set the position of table in the word doc to start from (0 is the 1st table)
table_position_range_start = 1
# Set the position of table in the word doc to end from
table_position_range_end = 4

# Open the Word document
doc = Document(word_doc)

# starting number for the CSV file names
csv_number = 0

# Iterate through tables from position range
for i, table in enumerate(doc.tables[table_position_range_start:table_position_range_end], start=table_position_range_start):
    print(i)
    # Increment the CSV number for each table
    csv_number += 1
    
    # Create a CSV file for the current table
    csv_filename = f"output/{word_doc_name}_table_{csv_number}.csv"
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        
        # Iterate through rows in the table
        for row in table.rows:
            # Extract data from each cell in the row
            row_data = [cell.text.strip() for cell in row.cells]
            # Write the row to the CSV file
            writer.writerow(row_data)
    
    print(f"Table {csv_number} from position {i} in the word doc, extracted and saved to '{csv_filename}'")