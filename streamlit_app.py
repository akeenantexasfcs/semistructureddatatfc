#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
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
    # Find the category row
    quality_row = None
    for idx, row in df_raw.iterrows():
        if 'Quality' in str(row.iloc[0]):
            quality_row = idx
            category = str(row.iloc[0]).strip()
            break
    
    if quality_row is None:
        raise ValueError("Could not find Quality category")

    # Initialize structure
    json_data = {
        "category": category,
        "entries": []
    }

    # Start processing after header row
    header_row = quality_row + 1
    data_start = header_row + 1

    # Variables to track current parent and subcategory
    current_parent_entry = None

    # Process rows
    for i in range(data_start, len(df_raw)):
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
                "terms": []
            }
            json_data["entries"].append(current_parent_entry)
            continue
        
        # Check for metrics
        percent_rr_used = safe_numeric_convert(row.iloc[2])
        if not any([percent_rr_used, 
                    safe_numeric_convert(row.iloc[3]), 
                    safe_numeric_convert(row.iloc[4])]):
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

        if current_parent_entry:
            current_parent_entry["terms"].append(entry_dict)
        else:
            json_data["entries"].append(entry_dict)

    return json_data

def create_excel_from_json(json_data):
    """Create formatted Excel from JSON data"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book

        # Ensure there is at least one visible sheet
        if not workbook.active:
            workbook.create_sheet(title="Sheet1")
        worksheet = workbook.active
        worksheet.title = "Sheet1"

        # Write category
        worksheet.append([json_data["category"]])
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
        worksheet.cell(row=1, column=1).font = Font(bold=True)

        # Header
        headers = [
            "Name/Term", "LGD", "% RR Used", "% AGG Used", "Used", 
            "Available", "Total Exposure", "% TE of RR", "% TE of AGG"
        ]
        worksheet.append(headers)
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=2, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        # Write entries
        current_row = 3
        for entry in json_data["entries"]:
            if "name" in entry:
                # Parent name row
                worksheet.append([entry["name"]])
                worksheet.merge_cells(start_row=current_row, start_column=1, 
                                      end_row=current_row, end_column=9)
                for col_num in range(1, 10):
                    worksheet.cell(row=current_row, column=col_num).fill = PatternFill(
                        start_color='FFFFCC', 
                        end_color='FFFFCC', 
                        fill_type='solid'
                    )
                current_row += 1

            # Sub rows
            for term in entry.get("terms", []):
                metrics = term["metrics"]
                row_data = [
                    term["term"],
                    term["lgd"],
                    metrics.get("percentRRUsed"),
                    metrics.get("percentAGGUsed"),
                    metrics.get("used"),
                    metrics.get("available"),
                    metrics.get("totalExposure"),
                    metrics.get("percentTERR"),
                    metrics.get("percentTEAGG")
                ]
                worksheet.append(row_data)
                current_row += 1

        # Adjust column widths
        for col_num in range(1, 10):
            worksheet.column_dimensions[get_column_letter(col_num)].width = 15

        # Ensure all sheets are visible
        for sheet in workbook.sheetnames:
            workbook[sheet].sheet_state = 'visible'

    output.seek(0)
    return output

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
                        df_raw = pd.read_excel(uploaded_file, 
                                               sheet_name=sheet_name, 
                                               header=None)

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

