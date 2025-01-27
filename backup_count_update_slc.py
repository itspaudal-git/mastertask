import gspread
from google.oauth2.service_account import Credentials

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

# Authenticate using the service account credentials
creds = Credentials.from_service_account_file('cred.json', scopes=scope)
client = gspread.authorize(creds)

# Open the sheets
sheet1_id = "1pUWVG9B9WUMJmcGUom4zN4VlG1YjjOi_InWY9zNfxdg"
# sheet2_id = "1z6iH4j6K8oOzVUEl35WUf9S1bpbY6bI0azqIkdjecRw"
sheet2_id = input("Enter the term number: ")

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

col1_idx_2 = 0  # Column 1 in Master Task List
col2_idx_2 = 1  # Column 2 in Master Task List
col4_idx_2 = 3  # Column 4 in Master Task List
col13_idx_2 = 12  # Column 13 in Master Task List (Count)

# Create a lookup dictionary from sheet1
lookup_dict = {(row[col1_idx_1], row[col2_idx_1], row[col4_idx_1]): row[col13_idx_1] for row in data1}

# Collect updates for batch_update
updates = []
for i, row in enumerate(data2):
    key = (row[col1_idx_2], row[col2_idx_2], row[col4_idx_2])
    if key in lookup_dict:
        updates.append({
            'range': f'M{i + 6}',  # M corresponds to the 13th column (zero-based index 12), start at row 6 in Master Task List
            'values': [[lookup_dict[key]]]
        })

# Debugging: Print updates to be made
for update in updates:
    print(update)

# Apply all updates in one batch
if updates:
    sheet2.batch_update(updates)

print("Update complete.")
