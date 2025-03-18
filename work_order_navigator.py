
import streamlit as st
import pandas as pd
import os
from fpdf import FPDF

# File paths
override_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Planner WO BOM Swaps\\Planners WO BOM Swaps List.xlsx"
work_orders_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\Work Orders.xlsx"
inventory_file = "C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\Inventory.xlsx"

# Load data with error handling
def load_excel(file_path):
    if not os.path.exists(file_path):
        st.error(f"Error: File {file_path} not found.")
        return pd.DataFrame()
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        st.error(f"Error loading {file_path}: {e}")
        return pd.DataFrame()

work_orders = load_excel(work_orders_file)
inventory = load_excel(inventory_file)
overrides = load_excel(override_file)

# Streamlit UI
st.title("Work Order Pick List")
work_order_id = st.text_input("Enter Work Order ID:")

if work_order_id:
    filtered_work_order = work_orders[work_orders['WORKORDER_ID'].astype(str) == work_order_id]
    if not filtered_work_order.empty:
        st.subheader("Work Order Details")
        st.write(filtered_work_order)
        
        pick_list = pd.read_excel("C:\\Users\\sdunna\\OneDrive - KBI Biopharma\\Documents - CMF-SC\\Pick List Files\\Pick List.xlsx")
        filtered_pick_list = pick_list[pick_list['WORKORDER_ID'].astype(str) == work_order_id]
        
        if not filtered_pick_list.empty:
            st.subheader("Pick List Data")
            st.dataframe(filtered_pick_list)
            
            def generate_pdf():
                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.set_font("Arial", style='B', size=16)
                pdf.cell(200, 10, f"Work Order ID: {work_order_id}", ln=True, align='C')
                pdf.ln(10)
                
                pdf.set_font("Arial", size=12)
                for col in filtered_work_order.columns:
                    pdf.cell(0, 10, f"{col}: {filtered_work_order.iloc[0][col]}", ln=True)
                pdf.ln(5)
                
                pdf.set_font("Arial", style='B', size=14)
                pdf.cell(200, 10, "Pick List Data", ln=True, align='C')
                pdf.ln(5)
                
                pdf.set_font("Arial", size=10)
                for _, row in filtered_pick_list.iterrows():
                    row_text = " | ".join([f"{col}: {row[col]}" for col in filtered_pick_list.columns])
                    pdf.multi_cell(0, 7, row_text)
                    pdf.ln(3)
                
                pdf_file = "pick_list.pdf"
                pdf.output(pdf_file)
                return pdf_file
            
            if st.button("Download PDF"):
                pdf_file = generate_pdf()
                with open(pdf_file, "rb") as f:
                    st.download_button("Click to Download", f, file_name="Pick_List.pdf", mime="application/pdf")
        else:
            st.warning("No pick list data found for this Work Order ID.")
    else:
        st.error("Work Order ID not found.")
