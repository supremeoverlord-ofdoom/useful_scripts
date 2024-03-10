## Commands to enter in terminal

1. Check python is installed
`python --version`

2. If python is installed then enter following commands (if not installed go install)
3. Create and activate virtual environment just so python packages don't cause nay issues with other packages in your laptop

`python -m venv venv`
`venv/Scripts/activate`

4. Install packages needed for scripts
`pip install -r requirements.txt`

## Which script does what

word_table_to_csv.py - for when I need to get 1 table from 1 word doc into 1 csv, can run it multiple times 

word_tables_to_multi_csv.py - for when I need multiple tables from a word doc into separate csv files (ie when they have different column names and can't be combined into 1)

multi_word_tables_combine_to_csv.py - for when I need to get the same table in the same position but from multiple different word docs and combined into 1 csv
(note, edit the column names in combined_tables.csv to match the desired column names)

## How to run the script

5. eg this will run this script
`python multi_word_tables_combine_to_csv.py`

6. when finished with you work deactivate the venv
`deactivate`

