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
col3_idx_1 = 2  # Column 3 in Saltlake
col4_idx_1 = 3  # Column 4 in Saltlake
col5_idx_1 = 4  # Column 5 in Saltlake (Column E)
col6_idx_1 = 5  # Column 6 in Saltlake (Column F)
col7_idx_1 = 6  # Column 7 in Saltlake (Column G)
col8_idx_1 = 7  # Column 8 in Saltlake (Column H)

col1_idx_2 = 0  # Column 1 in Master Task List
col2_idx_2 = 1  # Column 2 in Master Task List
col3_idx_2 = 2  # Column 3 in Master Task List
col4_idx_2 = 3  # Column 4 in Master Task List
col5_idx_2 = 4  # Column 5 in Master Task List (Column E)
col6_idx_2 = 6  # Column 7 in Master Task List (Column G)
col13_idx_2 = 12  # Column 13 in Master Task List (Column M)
col47_idx_2 = 46  # Column 47 in Master Task List (Column AU)

def convert_to_number(value):
    """Convert a string to a number if possible, otherwise return the original string."""
    try:
        return float(value)
    except ValueError:
        return value

def normalize_task_name(name):
    """Normalize task name by stripping and converting to lowercase."""
    return name.strip().lower()

def process_task(task_name):
    # Normalize task name for consistent matching
    normalized_task_name = normalize_task_name(task_name)

    # Clear rows A to D and G on Master Task List where Column E matches the task_name
    clear_updates = []
    for i, row in enumerate(data2):
        if normalize_task_name(row[col5_idx_2]) == normalized_task_name:
            clear_updates.append({
                'range': f'A{i + 6}:D{i + 6}',
                'values': [[''] * 4]
            })
            clear_updates.append({
                'range': f'G{i + 6}',
                'values': [['']]
            })

    # Apply clear updates in one batch
    if clear_updates:
        sheet2.batch_update(clear_updates)

    # Bring over the data from Saltlake
    updates = []
    saltlake_rows = [src_row for src_row in data1 if normalize_task_name(src_row[col8_idx_1]) == normalized_task_name]
    saltlake_index = 0  # Track the index for the Saltlake data

    for i, row in enumerate(data2):
        if normalize_task_name(row[col5_idx_2]) == normalized_task_name and not any(row[col1_idx_2:col4_idx_2 + 1]):
            if saltlake_index < len(saltlake_rows):
                src_row = saltlake_rows[saltlake_index]
                updates.append({
                    'range': f'A{i + 6}:D{i + 6}',
                    'values': [src_row[col1_idx_1:col4_idx_1 + 1]]
                })
                updates.append({
                    'range': f'M{i + 6}',
                    'values': [[convert_to_number(src_row[col5_idx_1])]]
                })
                updates.append({
                    'range': f'AU{i + 6}',
                    'values': [[convert_to_number(src_row[col7_idx_1])]]
                })
                updates.append({
                    'range': f'G{i + 6}',
                    'values': [[convert_to_number(src_row[col6_idx_1])]]
                })
                saltlake_index += 1
            else:
                print(f"No more matching data found in Saltlake for {task_name}.")

    # Apply all updates in one batch
    if updates:
        sheet2.batch_update(updates)
        print(f"{task_name} updates applied successfully.")
    else:
        print(f"No updates were made for {task_name}.")

    print(f"{task_name} processing complete.")

# List of tasks to process
tasks_to_process = [
    "Cup Portioning",
    "Liquid Sachet Depositing",
    "Dry Sachet Depositing",
    "Tray Portioning and Sealing",
    "Drain",
    "Kettle",
    "Batch Mix",
    "Sauce Mix",
    "Open",
    "Oven",
    "VCM",
    "Thaw",
    "Band Sealing"
]

# Process each task
for task in tasks_to_process:
    process_task(task)
