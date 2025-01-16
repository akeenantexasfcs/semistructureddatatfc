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
    """Convert semi-structured Excel to JSON format matching React component structure"""
    try:
        # Get PD category from first row
        category = str(df_raw.iloc[0, 0]).strip()
        
        # Initialize JSON structure
        json_data = {
            "category": category,
            "entries": []
        }
        
        current_parent = None
        i = 1  # Start after PD row
        
        while i < len(df_raw):
            row = df_raw.iloc[i]
            
            # Skip empty rows
            if row.isna().all():
                i += 1
                continue
            
            name_term = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # If we have numeric data in the row
            has_metrics = pd.notna(row.iloc[2])
            
            # Parent row detection (name but no LGD and no metrics)
            if name_term and not lgd and not has_metrics:
                current_parent = {
                    "name": name_term,
                    "term": "",
                    "lgd": "",
                    "metrics": {}
                }
                
                # Check if next row indicates this is a REVOLVER section
                if i + 1 < len(df_raw):
                    next_row = df_raw.iloc[i + 1]
                    next_term = str(next_row.iloc[0]).strip() if pd.notna(next_row.iloc[0]) else ""
                    if "REVOLVER" in next_term.upper():
                        current_parent["subCategory"] = "REVOLVER"
                        current_parent["entries"] = []
                        i += 1  # Skip the REVOLVER row
                
                json_data["entries"].append(current_parent)
                i += 1
                continue
            
            # Process data row
            if has_metrics:
                metrics = {
                    "percentRRUsed": safe_numeric_convert(row.iloc[2]),
                    "percentAGGUsed": safe_numeric_convert(row.iloc[3]),
                    "used": safe_numeric_convert(row.iloc[4]),
                    "available": safe_numeric_convert(row.iloc[5]),
                    "totalExposure": safe_numeric_convert(row.iloc[6]),
                    "percentTERR": safe_numeric_convert(row.iloc[7]),
                    "percentTEAGG": safe_numeric_convert(row.iloc[8])
                }
                
                # If we have a parent with subCategory
                if current_parent and "subCategory" in current_parent:
                    current_parent["entries"].append({
                        "term": name_term,
                        "lgd": lgd,
                        "metrics": metrics
                    })
                # If we have a regular parent
                elif current_parent and not current_parent["term"]:
                    current_parent.update({
                        "term": name_term,
                        "lgd": lgd,
                        "metrics": metrics
                    })
                # No parent, standalone entry
                else:
                    json_data["entries"].append({
                        "name": "",
                        "term": name_term,
                        "lgd": lgd,
                        "metrics": metrics
                    })
                    current_parent = None
            
            i += 1
        
        return json_data
        
    except Exception as e:
        raise Exception(f"Error processing Excel: {str(e)}")

def create_excel_from_json(json_data):
    """Create formatted Excel from JSON data"""
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            workbook = writer.book
            if 'Sheet' in workbook.sheetnames:
                workbook.remove(workbook['Sheet'])
            worksheet = workbook.create_sheet('Data')
            
            # Headers
            headers = ['Name/Term', 'LGD', '% RR Used', '% AGG Used', 'Used', 
                      'Available', 'Total Exposure', '% TE of RR', '% TE of AGG']
            
            # Write category
            worksheet.cell(row=1, column=1, value="PD")
            worksheet.cell(row=2, column=1, value=json_data["category"])
            
            # Write column headers
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=3, column=col, value=header)
                cell.font = Font(bold=True)
            
            # Styles
            yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            
            current_row = 4
            
            # Write data
            for entry in json_data["entries"]:
                if entry["name"]:
                    # Write parent name
                    cell = worksheet.cell(row=current_row, column=1, value=entry["name"])
                    cell.fill = yellow_fill
                    cell.font = Font(bold=True)
                    current_row += 1
                
                # Write entry data
                if "entries" in entry:  # Has sub-entries
                    for sub_entry in entry["entries"]:
                        worksheet.cell(row=current_row, column=1, value=f"    {sub_entry['term']}")
                        worksheet.cell(row=current_row, column=2, value=sub_entry["lgd"])
                        metrics = sub_entry["metrics"]
                        
                        # Write metrics
                        if metrics.get("percentRRUsed") is not None:
                            cell = worksheet.cell(row=current_row, column=3, value=metrics["percentRRUsed"] / 100)
                            cell.number_format = '0.00%'
                        if metrics.get("percentAGGUsed") is not None:
                            cell = worksheet.cell(row=current_row, column=4, value=metrics["percentAGGUsed"] / 100)
                            cell.number_format = '0.00%'
                        if metrics.get("used") is not None:
                            cell = worksheet.cell(row=current_row, column=5, value=metrics["used"])
                            cell.number_format = '#,##0'
                        if metrics.get("available") is not None:
                            cell = worksheet.cell(row=current_row, column=6, value=metrics["available"])
                            cell.number_format = '#,##0'
                        if metrics.get("totalExposure") is not None:
                            cell = worksheet.cell(row=current_row, column=7, value=metrics["totalExposure"])
                            cell.number_format = '#,##0'
                        if metrics.get("percentTERR") is not None:
                            cell = worksheet.cell(row=current_row, column=8, value=metrics["percentTERR"] / 100)
                            cell.number_format = '0.00%'
                        if metrics.get("percentTEAGG") is not None:
                            cell = worksheet.cell(row=current_row, column=9, value=metrics["percentTEAGG"] / 100)
                            cell.number_format = '0.00%'
                        
                        current_row += 1
                elif entry["term"]:  # Single entry
                    worksheet.cell(row=current_row, column=1, value=f"    {entry['term']}")
                    worksheet.cell(row=current_row, column=2, value=entry["lgd"])
                    metrics = entry["metrics"]
                    
                    # Write metrics
                    if metrics.get("percentRRUsed") is not None:
                        cell = worksheet.cell(row=current_row, column=3, value=metrics["percentRRUsed"] / 100)
                        cell.number_format = '0.00%'
                    if metrics.get("percentAGGUsed") is not None:
                        cell = worksheet.cell(row=current_row, column=4, value=metrics["percentAGGUsed"] / 100)
                        cell.number_format = '0.00%'
                    if metrics.get("used") is not None:
                        cell = worksheet.cell(row=current_row, column=5, value=metrics["used"])
                        cell.number_format = '#,##0'
                    if metrics.get("available") is not None:
                        cell = worksheet.cell(row=current_row, column=6, value=metrics["available"])
                        cell.number_format = '#,##0'
                    if metrics.get("totalExposure") is not None:
                        cell = worksheet.cell(row=current_row, column=7, value=metrics["totalExposure"])
                        cell.number_format = '#,##0'
                    if metrics.get("percentTERR") is not None:
                        cell = worksheet.cell(row=current_row, column=8, value=metrics["percentTERR"] / 100)
                        cell.number_format = '0.00%'
                    if metrics.get("percentTEAGG") is not None:
                        cell = worksheet.cell(row=current_row, column=9, value=metrics["percentTEAGG"] / 100)
                        cell.number_format = '0.00%'
                    
                    current_row += 1
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 10
            for i in range(3, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(i)].width = 15
            
            # Apply alignments
            for row in worksheet.iter_rows(min_row=3, max_row=current_row-1):
                for cell in row:
                    if cell.column == 1:  # Name/Term
                        cell.alignment = Alignment(horizontal='left')
                    elif cell.column == 2:  # LGD
                        cell.alignment = Alignment(horizontal='center')
                    else:  # Numeric columns
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
                        
                        # Convert to JSON structure
                        json_data = process_excel_to_json(df_raw)
                        
                        # Display JSON for verification
                        st.write(f"JSON structure for sheet '{sheet_name}':")
                        st.json(json_data)
                        
                        # Create formatted Excel
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

