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
        # Get PD from first row (usually in A2)
        pd_rating = str(df_raw.iloc[0, 0]).strip()
        
        # Initialize JSON structure
        json_data = {
            "category": pd_rating,
            "entries": []
        }
        
        current_parent = None
        current_entry = None
        
        # Process rows starting from row 1
        for idx, row in df_raw.iloc[1:].iterrows():
            # Skip completely empty rows
            if row.isna().all():
                continue
                
            name_term = str(row.iloc[0]).strip()
            lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Check if this is a parent row (company name)
            if pd.notna(name_term) and lgd == "":
                current_parent = {
                    "name": name_term,
                    "term": "",  # Will be filled by next row
                    "lgd": "",
                    "metrics": {},
                    "entries": []  # For multiple entries under same parent
                }
                json_data["entries"].append(current_parent)
                continue
            
            # If we have numeric data, this is a data row
            if pd.notna(row.iloc[2]):  # Check if % Used of RR exists
                metrics = {
                    "percentRRUsed": safe_numeric_convert(row.iloc[2]),
                    "percentAGGUsed": safe_numeric_convert(row.iloc[3]),
                    "used": safe_numeric_convert(row.iloc[4]),
                    "available": safe_numeric_convert(row.iloc[5]),
                    "totalExposure": safe_numeric_convert(row.iloc[6]),
                    "percentTERR": safe_numeric_convert(row.iloc[7]),
                    "percentTEAGG": safe_numeric_convert(row.iloc[8])
                }
                
                # If we have a parent, this is a child entry
                if current_parent:
                    if current_parent["term"] == "":  # First entry
                        current_parent["term"] = name_term
                        current_parent["lgd"] = lgd
                        current_parent["metrics"] = metrics
                    else:  # Additional entries
                        current_parent["entries"].append({
                            "term": name_term,
                            "lgd": lgd,
                            "metrics": metrics
                        })
                else:  # No parent, create new entry
                    json_data["entries"].append({
                        "name": "",
                        "term": name_term,
                        "lgd": lgd,
                        "metrics": metrics
                    })
        
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
            
            # Headers
            headers = ['Name/Term', 'LGD', '% RR Used', '% AGG Used', 'Used', 
                      'Available', 'Total Exposure', '% TE of RR', '% TE of AGG']
            
            # Write category in row 1
            worksheet.cell(row=1, column=1, value=json_data["category"])
            worksheet.cell(row=1, column=1).font = Font(bold=True)
            
            # Write headers in row 2
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=2, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
            # Styles
            yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Write data
            current_row = 3
            for entry in json_data["entries"]:
                # Write parent name if exists
                if entry["name"]:
                    cell = worksheet.cell(row=current_row, column=1, value=entry["name"])
                    cell.font = Font(bold=True)
                    cell.fill = yellow_fill
                    current_row += 1
                
                # Write main entry
                if entry["term"]:
                    worksheet.cell(row=current_row, column=1, value=entry["term"])
                    worksheet.cell(row=current_row, column=2, value=entry["lgd"])
                    metrics = entry["metrics"]
                    worksheet.cell(row=current_row, column=3, value=f"{metrics['percentRRUsed']:.2f}%" if metrics.get('percentRRUsed') else "")
                    worksheet.cell(row=current_row, column=4, value=f"{metrics['percentAGGUsed']:.2f}%" if metrics.get('percentAGGUsed') else "")
                    worksheet.cell(row=current_row, column=5, value=metrics.get('used', ""))
                    worksheet.cell(row=current_row, column=6, value=metrics.get('available', ""))
                    worksheet.cell(row=current_row, column=7, value=metrics.get('totalExposure', ""))
                    worksheet.cell(row=current_row, column=8, value=f"{metrics['percentTERR']:.2f}%" if metrics.get('percentTERR') else "")
                    worksheet.cell(row=current_row, column=9, value=f"{metrics['percentTEAGG']:.2f}%" if metrics.get('percentTEAGG') else "")
                    current_row += 1
                
                # Write additional entries if they exist
                for sub_entry in entry.get("entries", []):
                    worksheet.cell(row=current_row, column=1, value=sub_entry["term"])
                    worksheet.cell(row=current_row, column=2, value=sub_entry["lgd"])
                    metrics = sub_entry["metrics"]
                    worksheet.cell(row=current_row, column=3, value=f"{metrics['percentRRUsed']:.2f}%" if metrics.get('percentRRUsed') else "")
                    worksheet.cell(row=current_row, column=4, value=f"{metrics['percentAGGUsed']:.2f}%" if metrics.get('percentAGGUsed') else "")
                    worksheet.cell(row=current_row, column=5, value=metrics.get('used', ""))
                    worksheet.cell(row=current_row, column=6, value=metrics.get('available', ""))
                    worksheet.cell(row=current_row, column=7, value=metrics.get('totalExposure', ""))
                    worksheet.cell(row=current_row, column=8, value=f"{metrics['percentTERR']:.2f}%" if metrics.get('percentTERR') else "")
                    worksheet.cell(row=current_row, column=9, value=f"{metrics['percentTEAGG']:.2f}%" if metrics.get('percentTEAGG') else "")
                    current_row += 1
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 40  # Name/Term
            worksheet.column_dimensions['B'].width = 10  # LGD
            for i in range(3, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(i)].width = 15
            
            # Apply borders and alignment to data range
            for row in worksheet.iter_rows(min_row=2, max_row=current_row-1, min_col=1, max_col=len(headers)):
                for cell in row:
                    cell.border = thin_border
                    if isinstance(cell.value, str) and '%' in str(cell.value):
                        cell.alignment = Alignment(horizontal='right')
                    elif isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal='right')
                        cell.number_format = '#,##0'
                    elif cell.column == 2:  # LGD column
                        cell.alignment = Alignment(horizontal='center')
                    else:
                        cell.alignment = Alignment(horizontal='left')
        
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
                        
                        # Create clean Excel from JSON
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

