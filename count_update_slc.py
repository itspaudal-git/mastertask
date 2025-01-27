import gspread
from google.oauth2.service_account import Credentials
from drive_utils import find_sheet_id  # Import the find_sheet_id function

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

# Authenticate using the service account credentials
creds = Credentials.from_service_account_file('cred.json', scopes=scope)
client = gspread.authorize(creds)

# Define the parent folder ID and shared drive ID
parent_folder_id = '1fADilXve_5uGuJgvlXx81dACz9QVH1lN'
drive_id = '0ANQRPTqApmSQUk9PVA'

# Get user input for the term
user_term = input("Enter the term number: ")

# Find the sheet2_id based on the term
sheet2_id = find_sheet_id(user_term, parent_folder_id, drive_id)

if not sheet2_id:
    print("Sheet not found. Exiting...")
    exit(1)

# Open the sheets
sheet1_id = "1pUWVG9B9WUMJmcGUom4zN4VlG1YjjOi_InWY9zNfxdg"
sheet1 = client.open_by_key(sheet1_id).worksheet("Saltlake")
sheet2 = client.open_by_key(sheet2_id).worksheet("Master Task List")

# Get all values from the sheets
data1 = sheet1.get_all_values()
data2 = sheet2.get_all_values()

# Remove headers and start from the actual data rows
data1 = data1[1:]  # Start from row 2
data2 = data2[5:]  # Start from row 6

# Specify column indices (0-based)
col1_idx_1 = 0  # Column 1 in Saltlake
col2_idx_1 = 1  # Column 2 in Saltlake
col4_idx_1 = 3  # Column 4 in Saltlake
col13_idx_1 = 4  # Column 5 in Saltlake (Qty)
col56_idx_1 = 8  # Column 9 in Saltlake (Yield)

col1_idx_2 = 0  # Column 1 in Master Task List
col2_idx_2 = 1  # Column 2 in Master Task List
col4_idx_2 = 3  # Column 4 in Master Task List
col13_idx_2 = 12  # Column 13 in Master Task List (Count)
col56_idx_2 = 55  # Column 56 in Master Task List (Yield)

# Create lookup dictionaries from sheet1
lookup_dict_qty = {(row[col1_idx_1], row[col2_idx_1], row[col4_idx_1]): row[col13_idx_1] for row in data1}
lookup_dict_yield = {(row[col1_idx_1], row[col2_idx_1], row[col4_idx_1]): row[col56_idx_1] for row in data1}

# Collect updates for batch_update
updates = []
for i, row in enumerate(data2):
    key = (row[col1_idx_2], row[col2_idx_2], row[col4_idx_2])
    
    # Update Column 13 (Count) in Master Task List
    if key in lookup_dict_qty:
        try:
            numeric_value_qty = float(lookup_dict_qty[key])
            updates.append({'range': f'M{i + 6}', 'values': [[numeric_value_qty]]})
        except ValueError:
            print(f"Non-numeric value found for {key}: {lookup_dict_qty[key]}")
    
    # Update Column 56 (Yield) in Master Task List
    if key in lookup_dict_yield:
        try:
            numeric_value_yield = float(lookup_dict_yield[key])
            updates.append({'range': f'BD{i + 6}', 'values': [[numeric_value_yield]]})  # BD is column 56
        except ValueError:
            print(f"Non-numeric value found for {key}: {lookup_dict_yield[key]}")

# Debugging: Print updates to be made
for update in updates:
    print(update)

# Print total number of rows to be updated
print(f"Total number of rows to be updated: {len(updates)}")

# Apply all updates in one batch
if updates:
    sheet2.batch_update(updates)

print("Update complete.")