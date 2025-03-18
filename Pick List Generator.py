import pandas as pd
import os

# File paths
override_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Planner WO BOM Swaps\\Planners WO BOM Swaps List.xlsx"
work_orders_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\Work Orders.xlsx"
inventory_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\Inventory.xlsx"

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

work_orders = load_excel(work_orders_file)
inventory = load_excel(inventory_file)
overrides = load_excel(override_file)

# Filter inventory based on CUSTOM_DATA1(Location Category) values
valid_location_categories = ["CMF Warehouse", "CMF Warehouse - Cold", "W1", "W2", "W3", "W4"]
inventory = inventory[inventory['CUSTOM_DATA1'].isin(valid_location_categories)]

# Ensure necessary columns exist
def validate_columns(df, required_columns, name):
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        print(f"Error: Missing columns {missing} in {name}")
        return False
    return True

if not validate_columns(work_orders, ['SCHED_DATETIME', 'COMP_ITEMID', 'WORKORDER_ID', 'SITE_ID', 'QTY', 'SEQ_NUM', 'ORIG_CODE', 'PROJECT_NUMBER', 'PROD_BATCH_NUM'], 'work_orders') or \
   not validate_columns(inventory, ['ITEM_ID', 'QTYTOTAL', 'EXPDATE', 'SKIDID', 'LOTID', 'LOC_ID', 'ITEMDESC', 'SITE_ID', 'CUSTOM_DATA1'], 'inventory') or \
   not validate_columns(overrides, ['Swap Level', 'Project Number', 'Work Order ID', 'Production Batch Number', 'Original KBI Item Number', 'Substitute KBI Item Number'], 'overrides'):
    exit()

# Convert columns to correct types
work_orders['SCHED_DATETIME'] = pd.to_datetime(work_orders['SCHED_DATETIME'], errors='coerce')
work_orders['COMP_ITEMID'] = pd.to_numeric(work_orders['COMP_ITEMID'], errors='coerce').dropna().astype(int)
inventory['ITEM_ID'] = pd.to_numeric(inventory['ITEM_ID'], errors='coerce').dropna().astype(int)

# Sort data
work_orders.sort_values(by=['SCHED_DATETIME'], inplace=True)
inventory.sort_values(by=['ITEM_ID', 'EXPDATE'], inplace=True)

# Create inventory dictionary
inventory_dict = inventory.groupby(['SITE_ID', 'ITEM_ID']).apply(lambda x: x.to_dict(orient='records')).to_dict()

# Function to get substitute based on override Swap Levels
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
    custom_data = row['CUSTOM_DATA1']
    item_id = get_substitute(row['COMP_ITEMID'], workorder_id, project_number, batch_number)
    qty_needed = row['QTY']
    original_qty_needed = qty_needed
    
    site_options = [row['SITE_ID']]
    if row['SITE_ID'] == 2:
        site_options.insert(0, 5)

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
                "Project ID": project_number,
                "Production Batch Number": batch_number,
                "Custom Data": custom_data,
                "WORKORDER_ID": workorder_id,
                "ITEM_ID": item_id,
                "TOTAL_QTY_TO_PICK": original_qty_needed,
                "LOT_QTY_TO_PICK": qty_to_pick,
                "SKIDID": lot['SKIDID'],
                "LOTID": lot['LOTID'],
                "LOC_ID": lot['LOC_ID'],
                "Expiration Date": lot['EXPDATE'],
                "Item Description": lot['ITEMDESC'],
                "Stock UoM": lot['STOCK_UOM'],
                "SOURCE_SITE_ID": site_id,
                "TARGET_SITE_ID": row['SITE_ID'],
                "SCHED_DATETIME": row['SCHED_DATETIME'],
                "UNFULFILLED_QTY": original_qty_needed - total_allocated,
                "ALLOCATION_STATUS": "Fully Allocated" if total_allocated == original_qty_needed else ("Partially Allocated" if total_allocated > 0 else "Not Allocated")
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
