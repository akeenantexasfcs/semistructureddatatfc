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
    # First row has the category
    category = str(df_raw.iloc[0, 0]).strip()
    
    # Initialize structure
    json_data = {
        "category": category,
        "entries": []
    }
    
    # Variables to track current parent and subcategory
    current_parent_entry = None
    current_subcategory_entry = None
    
    # Process rows starting from row 2
    for i in range(1, len(df_raw)):
        row = df_raw.iloc[i]
        
        # Skip empty rows
        if row.isna().all():
            continue

        # Extract columns
        col_name_term = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        col_lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        
        # Check for parent row
        if col_name_term and not col_lgd:
            current_parent_entry = {
                "name": col_name_term,
            }
            json_data["entries"].append(current_parent_entry)
            current_subcategory_entry = None
            continue
        
        # Check for subcategory
        if "REVOLVER" in col_name_term.upper():
            current_subcategory_entry = {
                "name": current_parent_entry["name"] if current_parent_entry else "",
                "subCategory": col_name_term,
                "entries": []
            }
            json_data["entries"].append(current_subcategory_entry)
            continue
        
        # Check for metrics
        percent_rr_used = safe_numeric_convert(row.iloc[2])
        if not any([percent_rr_used, safe_numeric_convert(row.iloc[3]), safe_numeric_convert(row.iloc[4])]):
            continue
        
        # Build metrics
        metrics = {
            "percentRRUsed": percent_rr_used,
            "percentAGGUsed": safe_numeric_convert(row.iloc[3]),
            "used": safe_numeric_convert(row.iloc[4]),
            "available": safe_numeric_convert(row.iloc[5]),
            "totalExposure": safe_numeric_convert(row.iloc[6]),
            "percentTERR": safe_numeric_convert(row.iloc[7]),
            "percentTEAGG": safe_numeric_convert(row.iloc[8])
        }
        
        entry_dict = {
            "term": col_name_term,
            "lgd": col_lgd,
            "metrics": metrics
        }
        
        if current_subcategory_entry:
            current_subcategory_entry["entries"].append(entry_dict)
        elif current_parent_entry:
            current_parent_entry.update(entry_dict)
        else:
            json_data["entries"].append({
                "name": "",
                **entry_dict
            })
    
    return json_data

def create_excel_from_json(json_data):
    """Create formatted Excel from JSON data"""
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            workbook = writer.book
            worksheet = workbook.active  # Use active sheet
            worksheet.title = 'Sheet1'   # Rename it
            
            # Headers
            headers = [
                "Name/Term", "LGD", "% RR Used", "% AGG Used", "Used",
                "Available", "Total Exposure", "% TE of RR", "% TE of AGG"
            ]
            
            # Row 1: Category
            worksheet.cell(row=1, column=1, value=json_data["category"])
            worksheet.cell(row=1, column=1).font = Font(bold=True)
            
            # Row 2: Headers
            for col_idx, header in enumerate(headers, start=1):
                cell = worksheet.cell(row=2, column=col_idx, value=header)
                cell.font = Font(bold=True)
            
            # Styles
            yellow_fill = PatternFill(start_color='FFEB9C',
                                    end_color='FFEB9C',
                                    fill_type='solid')
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            current_row = 3
            
            def write_data_row(row_num, term, lgd, metrics, indent=0):
                # Term
                worksheet.cell(row=row_num, column=1, value=(" " * indent) + (term or ""))
                # LGD
                worksheet.cell(row=row_num, column=2, value=lgd)
                
                # Metrics
                if metrics.get("percentRRUsed") is not None:
                    cell = worksheet.cell(row=row_num, column=3, value=metrics["percentRRUsed"] / 100)
                    cell.number_format = '0.00%'
                if metrics.get("percentAGGUsed") is not None:
                    cell = worksheet.cell(row=row_num, column=4, value=metrics["percentAGGUsed"] / 100)
                    cell.number_format = '0.00%'
                if metrics.get("used") is not None:
                    cell = worksheet.cell(row=row_num, column=5, value=metrics["used"])
                    cell.number_format = '#,##0'
                if metrics.get("available") is not None:
                    cell = worksheet.cell(row=row_num, column=6, value=metrics["available"])
                    cell.number_format = '#,##0'
                if metrics.get("totalExposure") is not None:
                    cell = worksheet.cell(row=row_num, column=7, value=metrics["totalExposure"])
                    cell.number_format = '#,##0'
                if metrics.get("percentTERR") is not None:
                    cell = worksheet.cell(row=row_num, column=8, value=metrics["percentTERR"] / 100)
                    cell.number_format = '0.00%'
                if metrics.get("percentTEAGG") is not None:
                    cell = worksheet.cell(row=row_num, column=9, value=metrics["percentTEAGG"] / 100)
                    cell.number_format = '0.00%'
            
            for entry in json_data["entries"]:
                name = entry.get("name", "")
                sub_cat = entry.get("subCategory", "")
                
                if sub_cat:
                    # Write parent with subcategory
                    row_cell = worksheet.cell(row=current_row, column=1, value=name + " - " + sub_cat)
                    row_cell.fill = yellow_fill
                    current_row += 1
                    
                    # Write sub-entries
                    for sub_entry in entry["entries"]:
                        write_data_row(
                            current_row,
                            sub_entry.get("term"),
                            sub_entry.get("lgd"),
                            sub_entry.get("metrics", {}),
                            indent=2
                        )
                        current_row += 1
                else:
                    if "entries" in entry and isinstance(entry["entries"], list):
                        # Write parent name
                        if name:
                            parent_cell = worksheet.cell(row=current_row, column=1, value=name)
                            parent_cell.fill = yellow_fill
                            current_row += 1
                        
                        # Write sub-entries
                        for sub_entry in entry["entries"]:
                            write_data_row(
                                current_row,
                                sub_entry.get("term"),
                                sub_entry.get("lgd"),
                                sub_entry.get("metrics", {}),
                                indent=2
                            )
                            current_row += 1
                    else:
                        # Handle single entry
                        if name and not entry.get("term"):
                            parent_cell = worksheet.cell(row=current_row, column=1, value=name)
                            parent_cell.fill = yellow_fill
                            current_row += 1
                        
                        if entry.get("term"):
                            write_data_row(
                                current_row,
                                entry["term"],
                                entry.get("lgd", ""),
                                entry.get("metrics", {}),
                                indent=2
                            )
                            current_row += 1
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 10
            for col_idx in range(3, 10):
                worksheet.column_dimensions[get_column_letter(col_idx)].width = 15
            
            # Apply borders and alignment
            for row_cells in worksheet.iter_rows(min_row=2, max_row=current_row - 1, min_col=1, max_col=9):
                for cell in row_cells:
                    cell.border = thin_border
                    if cell.column in [3,4,5,6,7,8,9]:
                        cell.alignment = Alignment(horizontal='right')
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
                        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                        
                        # Convert to JSON
                        json_data = process_excel_to_json(df_raw)
                        
                        # Create Excel
                        excel_data = create_excel_from_json(json_data)
                        st.download_button(
                            label=f"Download {sheet_name}",
                            data=excel_data,
                            file_name=f"{sheet_name}_formatted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"Error processing sheet '{sheet_name}': {str(e)}")
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")

if __name__ == "__main__":
    main()

