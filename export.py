import xlwings as xw
import csv
import time
import json
import os
import string

def getcsd():
    return os.getcwd() + os.sep

# Connect to Excel
app = xw.apps.active
wb = app.books[0]
sheet = wb.sheets[0]

# JSON path
row_header_keypath = getcsd() + "keys.json"

# Initialize sets and mappings
pos0_set = set()
pos1_set = set()
local_row_name_to_index = {}
local_col_name_to_alpha_index = {}

# If keys.json doesn't exist, scan and build it
if not os.path.exists(row_header_keypath):
    for cell in sheet.used_range:
        formula = cell.formula
        if isinstance(formula, str) and formula.startswith('=BDP('):
            try:
                args = formula[5:-1].split(',')
                param0 = args[0].strip().strip('"')
                param1 = args[1].strip().strip('"')
                pos0_set.add(param0)
                pos1_set.add(param1)
            except Exception:
                continue

    pos0_list = sorted(list(pos0_set))
    pos1_list = sorted(list(pos1_set))

    # Map tickers to row numbers starting from 2
    for i, ticker in enumerate(pos0_list):
        local_row_name_to_index[ticker] = i + 2

    # Map fields to column letters starting from B
    for i, field in enumerate(pos1_list):
        col_letter = string.ascii_uppercase[i + 1]  # B, C, D...
        local_col_name_to_alpha_index[field] = col_letter

    with open(row_header_keypath, 'w') as f:
        json.dump({
            "pos0": pos0_list,
            "pos1": pos1_list
        }, f, indent=2)

    # Clear sheet and write formulas
    sheet.clear_contents()
    sheet.range("A1").expand().clear()  # Clears all used cells
    sheet.api.Cells.ClearFormats()      # Clears formatting
    sheet.api.Cells.ClearComments()     # Clears comments
    sheet.api.Cells.ClearHyperlinks()   # Clears hyperlinks
    sheet.api.Cells.ClearNotes()        
    # Write column headers
    for field, col in local_col_name_to_alpha_index.items():
        sheet.range(f"{col}1").value = field
    # Write row headers and formulas
    for ticker, row in local_row_name_to_index.items():
        sheet.range(f"A{row}").value = ticker
        for field, col in local_col_name_to_alpha_index.items():
            formula = f'=BDP("{ticker}","{field}")'
            sheet.range(f"{col}{row}").formula = formula

# Load keys
with open(row_header_keypath, 'r') as f:
    keys = json.load(f)
    pos0_list = keys["pos0"]
    pos1_list = keys["pos1"]

# Prepare output directory
os.makedirs("data", exist_ok=True)

def query_download(app, sec_freq, iterations):
    for i in range(iterations):
        timestamp = int(time.time())
        csv_path = f'data/W2EQ_{timestamp}.csv'

        with open(csv_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Row Header', 'Column Header', 'Value', 'Immediate Clause'])

            for ticker in pos0_list:
                row = local_row_name_to_index[ticker]
                for field in pos1_list:
                    col = local_col_name_to_alpha_index[field]
                    cell = sheet.range(f"{col}{row}")
                    value = cell.value
                    clause = 'BDP'
                    writer.writerow([ticker, field, value, clause])

        print(f"[{i+1}/{iterations}] CSV saved to: {csv_path}")
        time.sleep(sec_freq)

# Run the function
if __name__ == "__main__":
    query_download(app=app, sec_freq=1, iterations=5)
