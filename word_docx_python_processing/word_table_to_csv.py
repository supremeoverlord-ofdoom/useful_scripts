from docx import Document
import csv

## User input
# Set the position of the table in the word doc you want to extract (0 means the first table ect)
table_position = 1

# Open the Word document
doc = Document('input/example_doc.docx')

# Access the first table in the document
table = doc.tables[table_position]

# Create a CSV file to write the table data
with open('output/table_data.csv', 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)

    # Iterate through rows in the table
    for row in table.rows:
        # Extract data from each cell in the row
        row_data = [cell.text.strip() for cell in row.cells]
        # Write the row to the CSV file
        writer.writerow(row_data)

print("Table data extracted and saved to 'table_data.csv'")