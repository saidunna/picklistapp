import streamlit as st
import pandas as pd
from fpdf import FPDF
import math
from datetime import datetime
from collections import defaultdict

st.set_page_config(layout="wide", page_title="Work Order Pick List App", initial_sidebar_state="expanded")

# Load data
pick_list = pd.read_excel("/workspaces/picklistapp/Pick List.xlsx")

# Custom CSS for lighter theme and aesthetics
st.markdown(
    """
    <style>
    body {
        background-color: #f0f8ff; /* Light blue background */
        color: #333333; /* Darker text for better contrast */
    }

    .title {
        font-size: 32px;
        color: #4CAF50; /* Fresh green for the title */
        font-weight: bold;
    }

    .sidebar .sidebar-content {
        background-color: #eaf1f1; /* Light teal sidebar background */
        color: #333333; /* Dark text for readability */
    }

    .stButton > button {
        background-color: #4CAF50; /* Lighter green for buttons */
        color: white;
        font-size: 16px;
        padding: 12px;
        border-radius: 5px;
    }

    .stButton > button:hover {
        background-color: #45a049; /* Slightly darker green for hover effect */
    }

    .stSelectbox select, .stTextInput input {
        background-color: #ffffff; /* White background for input fields */
        color: #333333; /* Dark text for input fields */
        padding: 10px;
        border-radius: 5px;
        border: 1px solid #ddd; /* Light border color */
    }

    .stSelectbox select:focus, .stTextInput input:focus {
        border-color: #4CAF50; /* Highlight input border on focus */
    }

    .stDataFrame {
        background-color: #ffffff; /* White background for data table */
        color: #333333; /* Dark text for readability */
    }

    .stDataFrame thead th {
        background-color: #f4f4f4; /* Light grey header */
    }

    .stMarkdown {
        color: #333333;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Work Order Pick List")
st.markdown("<p class='title'>Work Order Pick List</p>", unsafe_allow_html=True)

# Sidebar inputs
st.sidebar.header("Work Order Search")
site_id = st.sidebar.text_input("Enter Site ID:")
if site_id:
    st.session_state['site_id'] = str(site_id)  # Ensure site_id is a string

# Step 2: Work Order ID Input (Only if Site ID is entered)
if 'site_id' in st.session_state:
    site_id_selected = str(st.session_state['site_id'])  # Ensure session site_id is treated as a string
    
    # Ensure the SITE_ID column is also treated as string
    pick_list['Source Site ID'] = pick_list['Source Site ID'].astype(str)
    
    # Check if the selected site_id exists in the work_orders DataFrame
    if site_id_selected in pick_list['Source Site ID'].values:
        # Filter work orders by the selected site_id
        filtered_work_orders = pick_list[pick_list['Source Site ID'] == site_id_selected]
        
        # Check if there are any work orders after filtering
        if not filtered_work_orders.empty:
            work_order_options = filtered_work_orders['Work Order ID'].astype(str).unique()
            work_order_id = st.sidebar.selectbox("Select Work Order ID:", options=work_order_options)
        else:
            st.warning(f"No work orders found for site ID {site_id_selected}.")
    else:
        st.warning(f"Site ID {site_id_selected} does not exist in the work orders.")
    
    if work_order_id:
        filtered_pick_list = pick_list[pick_list['Work Order ID'].astype(str) == work_order_id]
        if not filtered_pick_list.empty:
            # Display key work order details
            st.subheader("Work Order Details")
            col1, col2, col3, col4, col5, col6 = st.columns(6)
            col1.metric("Project ID", filtered_pick_list.iloc[0]['Project Number'])
            col2.metric("Production Item ID", filtered_pick_list.iloc[0]['Production ID'])
            col3.metric("Batch", filtered_pick_list.iloc[0]['Batch ID'])

            sched_datetime = pd.to_datetime(filtered_work_orders.iloc[0]['Scheduled Date'])
            sched_date_str = sched_datetime.strftime('%Y-%m-%d')
            col4.metric("Scheduled Date", sched_date_str)

            col5.metric("Custom Data", filtered_pick_list.iloc[0]['Custom Data'])
            col6.metric("BoM Custom Data", filtered_pick_list.iloc[0]['BoM Custom Data'])


            # Filter pick list data for the selected work order
            filtered_pick_list = pick_list[pick_list['Work Order ID'].astype(str) == work_order_id]
            if not filtered_pick_list.empty:
                st.subheader("Pick List Data")
                display_columns = ['Item ID', 'Total Qty to Pick', 'Location ID', 'Lot Qty to Pick', 'Lot ID', 'Expiration Date', 'Stock UoM', 'Item Description']
                st.dataframe(filtered_pick_list[display_columns])

                def generate_pdf():
                    pdf = FPDF()
                    # Add a Unicode font (DejaVu)
                    pdf.add_font('DejaVu', '', '/workspaces/picklistapp/DejaVuSans.ttf', uni=True)
                    pdf.add_font('DejaVu', 'B', '/workspaces/picklistapp/DejaVuSans-Bold.ttf', uni=True)
                    pdf.set_font('DejaVu', '', 10)

                    pdf.set_auto_page_break(auto=True, margin=15)
                    pdf.add_page()

                    # Title Section
                    pdf.set_font("Arial", 'B', 18)
                    pdf.set_fill_color(235, 235, 235)  # Light blue background
                    pdf.cell(0, 12, f"Work Order ID: {work_order_id}", ln=True, align='C', fill=True)
                    pdf.ln(8)

                    # Work Order Information - Two Column Layout
                    pdf.set_font("Arial", size=12)
                    
                    # Left Column
                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(45, 8, "Project ID:", border=0)
                    pdf.set_font("Arial", size=12)
                    pdf.cell(50, 8, f"{filtered_pick_list.iloc[0]['Project Number']}", border=0)
                    
                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(45, 8, "Production Item ID:", border=0)
                    pdf.set_font("Arial", size=12)
                    pdf.cell(50, 8, f"{filtered_pick_list.iloc[0]['Production ID']}", ln=True)

                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(45, 8, "Batch:", border=0)
                    pdf.set_font("Arial", size=12)
                    pdf.cell(50, 8, f"{filtered_pick_list.iloc[0]['Batch ID']}", border=0)

                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(45, 8, "Scheduled Date:", border=0)
                    pdf.set_font("Arial", size=12)
                    pdf.cell(50, 8, f"{filtered_pick_list.iloc[0]['Scheduled Date']}", ln=True)

                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(45, 8, "Custom Data:", border=0)
                    pdf.set_font("Arial", size=12)
                    pdf.cell(50, 8, f"{filtered_pick_list.iloc[0]['Custom Data']}", ln=True)

                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(45, 8, "BoM Custom Data:", border=0)
                    pdf.set_font("Arial", size=12)
                    pdf.cell(50, 8, f"{filtered_pick_list.iloc[0]['BoM Custom Data']}", ln=True)

                    pdf.ln(10)

                    # Pick List Header
                    pdf.set_font("Arial", 'B', 14)
                    pdf.cell(0, 10, "Pick List Data", ln=True, align='C', fill=True)
                    pdf.ln(4)

                    # Table Formatting
                    pdf.set_font("Arial", 'B', 10)

                    # Group rows by ITEM_ID
                    grouped_data = defaultdict(list)
                    for _, row in filtered_pick_list.iterrows():
                        item_id = row["Item ID"]
                        grouped_data[item_id].append(row)

                    # Start processing each item group
                    for item_id, item_rows in grouped_data.items():
                        # Extract item-specific details for the header
                        total_qty_to_pick = str(item_rows[0]["Total Qty to Pick"])
                        stock_uom = str(item_rows[0]["Stock UoM"])
                        item_description = str(item_rows[0]["Item Description"])
                        item_orig_status = str(item_rows[0]['Original/Substitute'])
                        allocation_status = str(item_rows[-1]["Allocation Status"])

                        # Display the item-specific details as headers, each on a new line
                        pdf.set_font("DejaVu", 'B', 10)  # Set bold for the name

                        pdf.cell(40, 10, f"Item ID: ", border=0, ln=False)  # Set width and avoid line break
                        pdf.set_font("DejaVu", '', 10)  # Set normal for the value
                        pdf.cell(0, 10, f"{item_id}", ln=True)  # Add line break after value

                        pdf.set_font("DejaVu", 'B', 10)  # Set bold for the name
                        pdf.cell(40, 10, f"Total Qty to Pick: ", border=0, ln=False)
                        pdf.set_font("DejaVu", '', 10)  # Set normal for the value
                        pdf.cell(0, 10, f"{item_orig_status}", ln=True)

                        pdf.set_font("DejaVu", 'B', 10)  # Set bold for the name
                        pdf.cell(40, 10, f"Total Qty to Pick: ", border=0, ln=False)
                        pdf.set_font("DejaVu", '', 10)  # Set normal for the value
                        pdf.cell(0, 10, f"{total_qty_to_pick}", ln=True)

                        pdf.set_font("DejaVu", 'B', 10)  # Set bold for the name
                        pdf.cell(40, 10, f"Stock UoM: ", border=0, ln=False)
                        pdf.set_font("DejaVu", '', 10)  # Set normal for the value
                        pdf.cell(0, 10, f"{stock_uom}", ln=True)

                        pdf.set_font("DejaVu", 'B', 10)  # Set bold for the name
                        pdf.cell(40, 10, f"Item Description: ", border=0, ln=False)
                        pdf.set_font("DejaVu", '', 10)  # Set normal for the value
                        pdf.cell(0, 10, f"{item_description}", ln=True)

                        pdf.set_font("DejaVu", 'B', 10)  # Set bold for the name
                        pdf.cell(40, 10, f"Allocation Status: ", border=0, ln=False)
                        pdf.set_font("DejaVu", '', 10)  # Set normal for the value
                        pdf.cell(0, 10, f"{allocation_status}", ln=True)                        

                        pdf.ln(6)  # Add some space between the item details and the sub-table


                        # Create the sub-table headers (with two columns)
                        sub_table_headers = ["LOC_ID", "LOT_QTY_TO_PICK", "LOTID", "Expiration Date"]
                        col_widths_sub = [15, 25, 25, 20]

                        # Print Sub-table Header (for location, lot qty, lot ID, expiration date) in two columns
                        pdf.set_font("DejaVu", '', 10)
                        pdf.cell(col_widths_sub[0] * 2, 10, sub_table_headers[0], border=1, align='C', fill=True)
                        pdf.cell(col_widths_sub[1] * 2, 10, sub_table_headers[1], border=1, align='C', fill=True)
                        pdf.cell(col_widths_sub[2] * 2, 10, sub_table_headers[2], border=1, align='C', fill=True)
                        pdf.cell(col_widths_sub[3] * 2, 10, sub_table_headers[3], border=1, align='C', fill=True)
                        pdf.ln(10)  # Move to the next row after headers

                        # Print each row in the sub-table for the current item
                        pdf.set_font("Arial", size=9)
                        for row in item_rows:
                            loc_id = str(row["LOC_ID"])
                            lot_qty_to_pick = str(row["LOT_QTY_TO_PICK"])
                            lotid = str(row["LOTID"])
                            expiration_date = row["Expiration Date"]
                            expiration_date = expiration_date.strftime('%m-%d-%Y')

                            # Print the data in the sub-table
                            pdf.cell(col_widths_sub[0] * 2, 10, loc_id, border=1, align='C')
                            pdf.cell(col_widths_sub[1] * 2, 10, lot_qty_to_pick, border=1, align='C')
                            pdf.cell(col_widths_sub[2] * 2, 10, lotid, border=1, align='C')
                            pdf.cell(col_widths_sub[3] * 2, 10, expiration_date, border=1, align='C')
                            pdf.ln(10)  # Move to the next row after this entry

                            # Ensure there's no overlap, check the Y position and add a page if needed
                            if pdf.get_y() > pdf.h - 40:  # 40 is a margin for the page bottom
                                pdf.add_page()

                        pdf.ln(10)  # Add some space between different items



                    # Save the file
                    pdf_file = f"WO_{work_order_id}.pdf"
                    pdf.output(pdf_file)
                    return pdf_file


                # Download button with a styled button
                if st.button("Download PDF", key="download", help="Click to download the PDF pick list", use_container_width=True):
                    pdf_file = generate_pdf()
                    with open(pdf_file, "rb") as f:
                        st.download_button("Click to Download", f, file_name=pdf_file, mime="application/pdf")
            else:
                st.warning("No pick list data found for this Work Order ID.")
        else:
            st.error("Work Order ID not found.")
