# J&T Express Invoice Date Extractor

Given a folder containing .xls/.xlsx files with J&T invoices, outputs a table of consolidated orders by company and month.

## Usage:
1. `git clone`
2. `pip install requirements.txt`
3. Change `PATH` to location of folder with invoices
4. Change `OUTPUT_PATH` to location of output file
5. `python main.py`

## Processing methodology

1. Get all sheets in file
1. Determine if sheet is valid 
1. Determine most common date from each file (cell C18)
1. For each sheet
    - Extract company name from B10
    - Extract date from C18
    - Extract order amount from E22
    - Ignore sheet if date doesn't match dominant date in file
1. Get order amount
1. Add to dictionary
1. Output dictionary to xlsx