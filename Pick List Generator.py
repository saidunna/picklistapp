import pandas as pd
import os
import pyodbc
import warnings

# Ignore the FutureWarning related to DataFrame concatenation
warnings.simplefilter(action='ignore', category=FutureWarning)

def load_query(filename):
    with open(filename, 'r') as file:
        return file.read()

# File paths for SQL queries    
work_orders_query_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\work_orders_query.sql"
inventory_query_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\inventory_query.sql"

work_orders_query = load_query(work_orders_query_file)
inventory_query = load_query(inventory_query_file)


# File path for SharePoint file
override_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Planner WO BOM Swaps\\Planners WO BOM Swaps List.xlsx"

# Load data with error handling
def load_excel(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File {file_path} not found.")
        return pd.DataFrame()
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        return pd.DataFrame()

# Set up the connection string with Windows Authentication
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=ddur-sql03\WWEST;' 
    r'DATABASE=FDW;'
    r'Trusted_Connection=yes;'
)

# Establish the connection
conn = pyodbc.connect(conn_str)

# Create a cursor to interact with the database
cursor = conn.cursor()

work_orders = pd.read_sql(work_orders_query, conn)
inventory = pd.read_sql(inventory_query, conn)
overrides = load_excel(override_file)

# CLose the connection
cursor.close()
conn.close()


# Filter inventory based on CUSTOM_DATA1(Location Category) values
valid_location_categories = ["CMF Warehouse", "CMF Warehouse - Cold", "W1", "W2", "W3", "W4"]
inventory = inventory[inventory['CUSTOM_DATA1'].isin(valid_location_categories)]

# Apply additional filtering logic
if 'BOM_CUSTOM_DATA1' in inventory.columns and 'CUSTOM_DATA1' in inventory.columns and 'ITEM_ID' in inventory.columns:
    inventory = inventory[
        (inventory['BOM_CUSTOM_DATA1'] != 'MFG Only') &  # Exclude 'MFG Only'
        ~((inventory['CUSTOM_DATA1'] == 'Downstream') & (inventory['ITEM_ID'].astype(str).str.startswith(('1', '7'))))  # Exclude Downstream + 1/7
    ]


# Convert columns to correct types
work_orders['SCHED_DATETIME'] = pd.to_datetime(work_orders['SCHED_DATETIME'], errors='coerce')
work_orders['COMP_ITEMID'] = pd.to_numeric(work_orders['COMP_ITEMID'], errors='coerce').dropna().astype(int)
inventory['ITEM_ID'] = pd.to_numeric(inventory['ITEM_ID'], errors='coerce').dropna().astype(int)

# Sort data
work_orders.sort_values(by=['SCHED_DATETIME'], inplace=True)
inventory.sort_values(by=['ITEM_ID', 'EXPDATE'], inplace=True)

# Create inventory dictionary
inventory_dict = inventory.groupby(['SITE_ID', 'ITEM_ID']).apply(lambda x: x.to_dict(orient='records')).to_dict()

# Function to get substitute based on override levels
def get_substitute(item_id, workorder_id, project_number, batch_number):
    override = overrides[(
        (overrides['Swap Level'] == 'Project') & (overrides['Project Number'] == project_number) |
        (overrides['Swap Level'] == 'Batch') & (overrides['Production Batch Number'] == batch_number) |
        (overrides['Swap Level'] == 'Work Order') & (overrides['Work Order ID'] == workorder_id)
    )]
    
    if not override.empty:
        substitute_row = override[override['Original KBI Item Number'] == item_id].drop_duplicates(subset=['Original KBI Item Number'])
        if not substitute_row.empty:
            return substitute_row.iloc[0]['Substitute KBI Item Number']
    return item_id

# Function to find substitute if original quantity is insufficient
def find_substitute(original_seq_num):
    substitute_row = work_orders[work_orders['SEQ_NUM'] == original_seq_num]
    if not substitute_row.empty:
        return substitute_row.iloc[0]['COMP_ITEMID']
    return None

# Allocate inventory
allocations = []

for _, row in work_orders.iterrows():
    if row['ORIG_CODE'] == 'S':
        continue

    workorder_id = row['WORKORDER_ID']
    project_number = row['PROJECT_NUMBER']
    batch_number = row['PROD_BATCH_NUM']
    prod_batch_id = row['PROD_ITEMID']
    orig_code = row['ORIG_CODE']
    custom_data = row['CUSTOM_DATA1']
    bom_custom_data = row['BOM_CUSTOM_DATA1']
    item_id = get_substitute(row['COMP_ITEMID'], workorder_id, project_number, batch_number)
    qty_needed = row['QTY']
    site_id = row['SITE_ID']
    original_qty_needed = qty_needed

    site_options = [site_id]
    if site_id == '2':
        site_options.insert(0, '5')

    total_allocated = 0
    for site_id in site_options:
        if (site_id, item_id) not in inventory_dict:
            continue

        lot_list = inventory_dict[(site_id, item_id)]
        new_lot_list = []

        for lot in lot_list:
            if qty_needed <= 0:
                new_lot_list.append(lot)
                continue

            qty_available = lot['QTYTOTAL']
            if qty_available <= 0:
                continue

            qty_to_pick = min(qty_needed, qty_available)
            total_allocated += qty_to_pick

            allocations.append({
                "Project Number": project_number,
                "Batch ID": batch_number,
                "Production ID": prod_batch_id,
                "Custom Data": custom_data,
                "BoM Custom Data": bom_custom_data,
                "Work Order ID": workorder_id,
                "Item ID": item_id,
                "Original/Substitute": orig_code,
                "Total Qty to Pick": original_qty_needed,
                "Lot Qty to Pick": qty_to_pick,
                "Lot ID": lot['LOTID'],
                "Location ID": lot['LOC_ID'],
                "Expiration Date": lot['EXPDATE'],
                "Item Description": lot['ITEMDESC'],
                "Stock UoM": lot['STOCK_UOM'],
                "Source Site ID": site_id,
                "Target Site ID": row['SITE_ID'],
                "Scheduled Date": row['SCHED_DATETIME'],
                "Unfulfilled Qty": original_qty_needed - total_allocated,
                "Allocation Status": "Fully Allocated" if total_allocated == original_qty_needed else ("Partially Allocated" if total_allocated > 0 else "Not Allocated")
            })

            lot['QTYTOTAL'] -= qty_to_pick
            qty_needed -= qty_to_pick

            if lot['QTYTOTAL'] > 0:
                new_lot_list.append(lot)

        inventory_dict[(site_id, item_id)] = new_lot_list
        if qty_needed <= 0:
            break

    if qty_needed > 0:
        substitute_item = find_substitute(row['SEQ_NUM'])
        if substitute_item:
            new_row = row.copy()
            new_row['COMP_ITEMID'] = substitute_item
            new_row['ORIG_CODE'] = 'S'
            work_orders = pd.concat([work_orders, new_row.to_frame().T], ignore_index=True)


# Convert allocations to DataFrame and save
pick_list = pd.DataFrame(allocations)
pick_list.to_excel("C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\Pick List.xlsx", index=False)
