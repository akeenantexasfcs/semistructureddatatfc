#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import io
import json
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

def safe_numeric_convert(value):
    """Safely convert a value to numeric, returning None if not possible"""
    if pd.isna(value):
        return None
    try:
        if isinstance(value, str):
            value = value.replace(',', '').replace('%', '')
        return float(value)
    except (ValueError, TypeError):
        return None

def process_excel_to_json(df_raw):
    """Convert semi-structured Excel to JSON format"""
    try:
        # Get PD Rating
        pd_rating = str(df_raw.iloc[0, 0]).strip()
        
        # Initialize JSON structure
        json_data = {
            "category": pd_rating,
            "entries": []
        }
        
        i = 1  # Start after PD row
        while i < len(df_raw):
            row = df_raw.iloc[i]
            
            # Skip empty rows
            if row.isna().all():
                i += 1
                continue
            
            name_term = str(row.iloc[0]).strip()
            lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Process row data if it has numeric values
            metrics = None
            if pd.notna(row.iloc[2]):  # Check if % RR Used exists
                metrics = {
                    "percentRRUsed": safe_numeric_convert(row.iloc[2]),
                    "percentAGGUsed": safe_numeric_convert(row.iloc[3]),
                    "used": safe_numeric_convert(row.iloc[4]),
                    "available": safe_numeric_convert(row.iloc[5]),
                    "totalExposure": safe_numeric_convert(row.iloc[6]),
                    "percentTERR": safe_numeric_convert(row.iloc[7]),
                    "percentTEAGG": safe_numeric_convert(row.iloc[8])
                }
            
            # If it's a parent row (no LGD)
            if name_term and not lgd:
                json_data["entries"].append({
                    "name": name_term,
                    "entries": []
                })
            # If it's a data row
            elif metrics:
                entry = {
                    "term": name_term,
                    "lgd": lgd,
                    "metrics": metrics
                }
                
                # Add to last parent if exists, otherwise create new parent
                if json_data["entries"] and json_data["entries"][-1].get("name"):
                    json_data["entries"][-1]["entries"].append(entry)
                else:
                    json_data["entries"].append({
                        "name": "",
                        "entries": [entry]
                    })
            
            i += 1
        
        return json_data
    except Exception as e:
        raise Exception(f"Error converting to JSON: {str(e)}")

def create_excel_from_json(json_data):
    """Create clean tabular Excel from JSON data"""
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            workbook = writer.book
            worksheet = workbook.create_sheet('Sheet1')
            
            # Write PD in row 1
            worksheet.cell(row=1, column=1, value="PD")
            worksheet.cell(row=1, column=1).font = Font(bold=True)
            
            # Headers in row 2
            headers = ['Name/Term', 'LGD', '% RR Used', '% AGG Used', 'Used', 
                      'Available', 'Total Exposure', '% TE of RR', '% TE of AGG']
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=2, column=col, value=header)
                cell.font = Font(bold=True)
            
            # Write category in row 3
            cell = worksheet.cell(row=3, column=1, value=json_data["category"])
            cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            
            # Define styles
            yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            current_row = 4
            
            # Write data
            for parent_entry in json_data["entries"]:
                # Write parent name if it exists
                if parent_entry["name"]:
                    cell = worksheet.cell(row=current_row, column=1, value=parent_entry["name"])
                    cell.fill = yellow_fill
                    current_row += 1
                
                # Write child entries
                if "entries" in parent_entry:
                    for sub_entry in parent_entry["entries"]:
                        # Write term with indentation if there's a parent
                        indent = "  " if parent_entry["name"] else ""
                        worksheet.cell(row=current_row, column=1, value=indent + sub_entry["term"])
                        worksheet.cell(row=current_row, column=2, value=sub_entry["lgd"])
                        
                        # Write metrics
                        metrics = sub_entry["metrics"]
                        
                        # Percentages with proper formatting
                        for col, (key, format_str) in enumerate([
                            ("percentRRUsed", "0.00%"),
                            ("percentAGGUsed", "0.00%")
                        ], 3):
                            if metrics.get(key) is not None:
                                cell = worksheet.cell(row=current_row, column=col, value=metrics[key] / 100)
                                cell.number_format = format_str
                                cell.alignment = Alignment(horizontal='right')
                        
                        # Numbers with comma formatting
                        for col, key in enumerate(['used', 'available', 'totalExposure'], 5):
                            if metrics.get(key) is not None:
                                cell = worksheet.cell(row=current_row, column=col, value=metrics[key])
                                cell.number_format = '#,##0'
                                cell.alignment = Alignment(horizontal='right')
                        
                        # TE percentages
                        for col, key in enumerate(['percentTERR', 'percentTEAGG'], 8):
                            if metrics.get(key) is not None:
                                cell = worksheet.cell(row=current_row, column=col, value=metrics[key] / 100)
                                cell.number_format = '0.00%'
                                cell.alignment = Alignment(horizontal='right')
                        
                        current_row += 1
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 40  # Name/Term
            worksheet.column_dimensions['B'].width = 10  # LGD
            for i in range(3, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(i)].width = 15
            
            # Apply borders and default alignment
            for row in worksheet.iter_rows(min_row=2, max_row=current_row-1):
                for cell in row:
                    cell.border = thin_border
                    if cell.column == 2:  # LGD column
                        cell.alignment = Alignment(horizontal='center')
                    elif cell.column > 2:  # Numeric columns
                        cell.alignment = Alignment(horizontal='right')
        
        output.seek(0)
        return output
        
    except Exception as e:
        raise Exception(f"Error creating Excel file: {str(e)}")

def main():
    st.title("Excel PD Sheet Processor")
    st.write("Upload an Excel workbook with PD sheets to process.")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.info(f"Found {len(sheet_names)} sheets in the workbook")
            
            # Allow sheet selection
            selected_sheets = st.multiselect(
                "Select sheets to process",
                sheet_names,
                default=sheet_names[0] if sheet_names else None
            )
            
            if st.button("Process Selected Sheets"):
                for sheet_name in selected_sheets:
                    try:
                        # Read sheet without header
                        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                        
                        # Convert to JSON structure
                        json_data = process_excel_to_json(df_raw)
                        
                        # Display JSON for verification
                        st.write(f"JSON structure for sheet '{sheet_name}':")
                        st.json(json_data)
                        
                        # Create formatted Excel from JSON
                        excel_data = create_excel_from_json(json_data)
                        
                        # Provide download button
                        st.download_button(
                            label=f"Download {sheet_name}",
                            data=excel_data,
                            file_name=f"{sheet_name}_formatted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"Error processing sheet '{sheet_name}': {str(e)}")
                        st.write(f"Please ensure sheet '{sheet_name}' follows the expected format.")
                        
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            st.write("Please ensure you've uploaded a valid Excel file (.xlsx or .xls).")

if __name__ == "__main__":
    main()

