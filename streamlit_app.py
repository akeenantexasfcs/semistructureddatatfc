#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import json
import io
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def process_pd_sheet(df):
    """Convert Excel sheet to structured JSON format"""
    # Remove empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Get the PD category from the first few rows
    pd_category = df.iloc[1, 0]  # Usually in second row, first column
    
    structured_data = {
        "category": pd_category,
        "entries": []
    }
    
    current_entry = None
    i = 2  # Start after headers and PD category
    
    while i < len(df):
        row = df.iloc[i]
        
        # Skip empty rows
        if pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]):
            i += 1
            continue
            
        # Check if this is a main entry (highlighted in yellow in Excel)
        if not pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]):
            if current_entry:
                structured_data["entries"].append(current_entry)
            
            current_entry = {
                "name": row.iloc[0],
                "term": None,
                "lgd": None,
                "metrics": None,
                "sub_entries": []
            }
            
        # Check if this is a term/metrics row
        elif not pd.isna(row.iloc[1]):  # Has LGD value
            metrics = {
                "percentRRUsed": float(row.iloc[2]) if not pd.isna(row.iloc[2]) else None,
                "percentAGGUsed": float(row.iloc[3]) if not pd.isna(row.iloc[3]) else None,
                "used": float(row.iloc[4]) if not pd.isna(row.iloc[4]) else None,
                "available": float(row.iloc[5]) if not pd.isna(row.iloc[5]) else None,
                "totalExposure": float(row.iloc[6]) if not pd.isna(row.iloc[6]) else None,
                "percentTERR": float(row.iloc[7]) if not pd.isna(row.iloc[7]) else None,
                "percentTEAGG": float(row.iloc[8]) if not pd.isna(row.iloc[8]) else None
            }
            
            if current_entry and current_entry["metrics"] is None:
                current_entry["term"] = row.iloc[0] if not pd.isna(row.iloc[0]) else None
                current_entry["lgd"] = row.iloc[1]
                current_entry["metrics"] = metrics
            else:
                sub_entry = {
                    "term": row.iloc[0] if not pd.isna(row.iloc[0]) else None,
                    "lgd": row.iloc[1],
                    "metrics": metrics
                }
                if current_entry:
                    current_entry["sub_entries"].append(sub_entry)
        
        i += 1
    
    # Add the last entry if exists
    if current_entry:
        structured_data["entries"].append(current_entry)
    
    return structured_data

def create_styled_excel(processed_data):
    """Create Excel file with matching styles from the screenshot"""
    # [Your existing create_styled_excel function code here]
    # Keep all the code from your existing create_styled_excel function

def main():
    st.title("Excel PD Sheet Processor")
    st.write("Upload an Excel workbook with PD sheets to process.")
    
    # File uploader for Excel files
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            # Read all sheets from the Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.info(f"Found {len(sheet_names)} sheets in the workbook")
            
            # Allow user to select which sheets to process
            selected_sheets = st.multiselect(
                "Select sheets to process",
                sheet_names,
                default=sheet_names
            )
            
            if st.button("Process Selected Sheets"):
                processed_data = {}
                
                for sheet_name in selected_sheets:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                    processed_data[sheet_name] = process_pd_sheet(df)
                
                # Generate the styled Excel output
                excel_data = create_styled_excel(processed_data)
                
                # Provide download button
                st.download_button(
                    label="Download Formatted Excel",
                    data=excel_data,
                    file_name="formatted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Show preview of processed data
                st.subheader("Preview of Processed Data")
                for sheet_name, data in processed_data.items():
                    with st.expander(f"Sheet: {sheet_name}"):
                        st.write(data['category'])
                        for entry in data['entries']:
                            st.write(f"- {entry['name']}")
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

if __name__ == "__main__":
    main()

