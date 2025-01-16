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

def excel_to_json(df_raw):
    """Convert semi-structured Excel to JSON format"""
    try:
        # Get PD Rating from first row
        pd_rating = str(df_raw.iloc[0, 0]).strip()
        
        # Initialize JSON structure
        json_data = {
            "category": pd_rating,
            "entries": []
        }
        
        current_parent = None
        i = 2  # Skip header rows
        
        while i < len(df_raw):
            row = df_raw.iloc[i]
            
            # Skip empty rows
            if row.isna().all():
                i += 1
                continue
            
            name_term = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Skip subtotal rows
            if name_term.lower().startswith('sub total'):
                i += 1
                continue
            
            # Check if this is a parent row (company name)
            if name_term and not lgd and not row.iloc[2:].notna().any():
                current_parent = {
                    "name": name_term,
                    "entries": []
                }
                json_data["entries"].append(current_parent)
                i += 1
                continue
            
            # Process term row with data
            if pd.notna(row.iloc[2]):  # Has metrics
                metrics = {
                    "percentRRUsed": safe_numeric_convert(row.iloc[2]),
                    "percentAGGUsed": safe_numeric_convert(row.iloc[3]),
                    "used": safe_numeric_convert(row.iloc[4]),
                    "available": safe_numeric_convert(row.iloc[5]),
                    "totalExposure": safe_numeric_convert(row.iloc[6]),
                    "percentTERR": safe_numeric_convert(row.iloc[7]),
                    "percentTEAGG": safe_numeric_convert(row.iloc[8])
                }
                
                entry = {
                    "term": name_term,
                    "lgd": lgd,
                    "metrics": metrics
                }
                
                if current_parent:
                    current_parent["entries"].append(entry)
                else:
                    # Standalone entry if no parent
                    json_data["entries"].append({
                        "name": "",
                        "entries": [entry]
                    })
            
            i += 1
        
        # Remove any entries without data
        json_data["entries"] = [
            entry for entry in json_data["entries"]
            if entry["entries"]  # Only keep entries that have child entries
        ]
        
        return json_data
            
            # Data row
            if lgd:
                metrics = {
                    "percentRRUsed": safe_numeric_convert(row.iloc[2]),
                    "percentAGGUsed": safe_numeric_convert(row.iloc[3]),
                    "used": safe_numeric_convert(row.iloc[4]),
                    "available": safe_numeric_convert(row.iloc[5]),
                    "totalExposure": safe_numeric_convert(row.iloc[6]),
                    "percentTERR": safe_numeric_convert(row.iloc[7]),
                    "percentTEAGG": safe_numeric_convert(row.iloc[8])
                }
                
                # If we have a parent, add to its entries
                if current_entry:
                    if not current_entry["term"]:
                        current_entry["term"] = name_term
                        current_entry["lgd"] = lgd
                        current_entry["metrics"] = metrics
                    else:
                        current_entry["entries"].append({
                            "term": name_term,
                            "lgd": lgd,
                            "metrics": metrics
                        })
                else:
                    # Standalone entry
                    json_data["entries"].append({
                        "name": "",
                        "term": name_term,
                        "lgd": lgd,
                        "metrics": metrics,
                        "entries": []
                    })
        
        return json_data
        
    except Exception as e:
        raise Exception(f"Error converting to JSON: {str(e)}")

def json_to_excel(json_data):
    """Create formatted Excel from JSON data"""
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            workbook = writer.book
            worksheet = workbook.create_sheet('Sheet1')
            
            # Write headers
            headers = ['Name/Term', 'LGD', '% RR Used', '% AGG Used', 'Used', 
                      'Available', 'Total Exposure', '% TE of RR', '% TE of AGG']
            
            # Write PD header
            cell = worksheet.cell(row=1, column=1, value='PD')
            cell.font = Font(bold=True)
            
            # Write column headers
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=2, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
            # Write PD category
            cell = worksheet.cell(row=3, column=1, value=json_data["category"])
            cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            
            # Styles
            yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            current_row = 4
            
            # Write entries
            for entry in json_data["entries"]:
                # Write parent name if exists
                if entry["name"]:
                    cell = worksheet.cell(row=current_row, column=1, value=entry["name"])
                    cell.fill = yellow_fill
                    cell.font = Font(bold=True)
                    current_row += 1
                
                # Write main entry or first entry
                if entry["term"]:
                    # Term
                    worksheet.cell(row=current_row, column=1, value=entry["term"])
                    
                    # LGD (center-aligned)
                    lgd_cell = worksheet.cell(row=current_row, column=2, value=entry["lgd"])
                    lgd_cell.alignment = Alignment(horizontal='center')
                    
                    # Metrics
                    metrics = entry["metrics"]
                    # Percentages
                    for col, (key, label) in enumerate([
                        ("percentRRUsed", 3),
                        ("percentAGGUsed", 4),
                        ("percentTERR", 8),
                        ("percentTEAGG", 9)
                    ]):
                        if metrics.get(key):
                            cell = worksheet.cell(row=current_row, column=label, 
                                               value=metrics[key] / 100)
                            cell.number_format = '0.00%'
                    
                    # Numbers
                    for col, key in enumerate(['used', 'available', 'totalExposure'], 5):
                        if metrics.get(key):
                            cell = worksheet.cell(row=current_row, column=col, 
                                               value=metrics[key])
                            cell.number_format = '#,##0'
                    
                    current_row += 1
                
                # Write additional entries
                for sub_entry in entry.get("entries", []):
                    worksheet.cell(row=current_row, column=1, value=sub_entry["term"])
                    
                    lgd_cell = worksheet.cell(row=current_row, column=2, value=sub_entry["lgd"])
                    lgd_cell.alignment = Alignment(horizontal='center')
                    
                    metrics = sub_entry["metrics"]
                    # Percentages
                    for col, (key, label) in enumerate([
                        ("percentRRUsed", 3),
                        ("percentAGGUsed", 4),
                        ("percentTERR", 8),
                        ("percentTEAGG", 9)
                    ]):
                        if metrics.get(key):
                            cell = worksheet.cell(row=current_row, column=label, 
                                               value=metrics[key] / 100)
                            cell.number_format = '0.00%'
                    
                    # Numbers
                    for col, key in enumerate(['used', 'available', 'totalExposure'], 5):
                        if metrics.get(key):
                            cell = worksheet.cell(row=current_row, column=col, 
                                               value=metrics[key])
                            cell.number_format = '#,##0'
                    
                    current_row += 1
            
            # Set column widths and final formatting
            worksheet.column_dimensions['A'].width = 40  # Name/Term
            worksheet.column_dimensions['B'].width = 10  # LGD
            for i in range(3, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(i)].width = 15
                
            # Apply borders and alignment to data range
            for row in worksheet.iter_rows(min_row=2, max_row=current_row-1):
                for cell in row:
                    cell.border = thin_border
                    if cell.column > 2:  # Columns after LGD
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
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.info(f"Found {len(sheet_names)} sheets in the workbook")
            
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
                        
                        # Convert to JSON
                        json_data = excel_to_json(df_raw)
                        
                        # Show JSON structure (for verification)
                        st.write(f"JSON structure for sheet '{sheet_name}':")
                        st.json(json_data)
                        
                        # Convert back to Excel
                        excel_data = json_to_excel(json_data)
                        
                        # Provide download button
                        st.download_button(
                            label=f"Download {sheet_name}",
                            data=excel_data,
                            file_name=f"{sheet_name}_formatted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"Error processing sheet '{sheet_name}': {str(e)}")
                        st.write("Please ensure the sheet follows the expected format.")
                        
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            st.write("Please ensure you've uploaded a valid Excel file (.xlsx or .xls).")

if __name__ == "__main__":
    main()

