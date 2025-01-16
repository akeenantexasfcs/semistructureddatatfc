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
        
        current_parent = None
        i = 1  # Start after PD row
        
        while i < len(df_raw):
            row = df_raw.iloc[i]
            
            # Skip empty rows
            if row.isna().all():
                i += 1
                continue
            
            name_term = str(row.iloc[0]).strip()
            lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Check if this is a parent row (has name but no LGD)
            if pd.notna(name_term) and (pd.isna(row.iloc[1]) or str(row.iloc[1]).strip() == ""):
                current_parent = {
                    "name": name_term,
                    "entries": []
                }
                json_data["entries"].append(current_parent)
                i += 1
                continue
            
            # If we have numeric data, this is a data row
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
                
                entry = {
                    "term": name_term,
                    "lgd": lgd,
                    "metrics": metrics
                }
                
                if current_parent:
                    current_parent["entries"].append(entry)
                else:
                    # No parent, create new entry
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
                if entry["name"]:
                    # Write parent name
                    cell = worksheet.cell(row=current_row, column=1, value=entry["name"])
                    cell.fill = yellow_fill
                    current_row += 1
                
                # Write all sub-entries
                for sub_entry in entry["entries"]:
                    # Indent term if there's a parent
                    term_cell = worksheet.cell(row=current_row, column=1, 
                                            value=("  " if entry["name"] else "") + sub_entry["term"])
                    term_cell.alignment = Alignment(horizontal='left')
                    
                    # LGD
                    lgd_cell = worksheet.cell(row=current_row, column=2, value=sub_entry["lgd"])
                    lgd_cell.alignment = Alignment(horizontal='center')
                    
                    metrics = sub_entry["metrics"]
                    
                    # Percentages
                    if metrics["percentRRUsed"] is not None:
                        cell = worksheet.cell(row=current_row, column=3, value=metrics["percentRRUsed"] / 100)
                        cell.number_format = '0.00%'
                    if metrics["percentAGGUsed"] is not None:
                        cell = worksheet.cell(row=current_row, column=4, value=metrics["percentAGGUsed"] / 100)
                        cell.number_format = '0.00%'
                    
                    # Numbers
                    for col, key in enumerate(['used', 'available', 'totalExposure'], 5):
                        if metrics[key] is not None:
                            cell = worksheet.cell(row=current_row, column=col, value=metrics[key])
                            cell.number_format = '#,##0'
                            cell.alignment = Alignment(horizontal='right')
                    
                    # TE Percentages
                    if metrics["percentTERR"] is not None:
                        cell = worksheet.cell(row=current_row, column=8, value=metrics["percentTERR"] / 100)
                        cell.number_format = '0.00%'
                    if metrics["percentTEAGG"] is not None:
                        cell = worksheet.cell(row=current_row, column=9, value=metrics["percentTEAGG"] / 100)
                        cell.number_format = '0.00%'
                    
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
                    if cell.column > 2:  # Columns after LGD
                        cell.alignment = Alignment(horizontal='right')
        
        output.seek(0)
        return output
        
    except Exception as e:
        raise Exception(f"Error creating Excel file: {str(e)}")
                
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

