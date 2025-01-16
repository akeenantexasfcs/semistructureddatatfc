#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

def safe_numeric_convert(value):
    """Safely convert a value to numeric, returning None if not possible"""
    if pd.isna(value):
        return None
    try:
        if isinstance(value, str):
            # Remove commas and % signs, handle scientific notation
            value = value.replace(',', '').replace('%', '')
            if 'E' in value.upper():
                # Handle scientific notation
                return float(value)
        return float(value)
    except (ValueError, TypeError):
        return None

def process_pd_data(df_raw):
    """Process the PD sheet data"""
    try:
        # Get PD Rating (first non-empty value in column A)
        pd_rows = df_raw[df_raw.iloc[:, 0].notna()]
        pd_rating = None
        for idx, row in pd_rows.iterrows():
            if 'quality' in str(row.iloc[0]).lower():
                pd_rating = str(row.iloc[0]).strip()
                break
        
        if not pd_rating:
            raise ValueError("Could not find PD rating in the sheet")

        # Initialize lists for storing data
        rows = []
        current_parent = None
        
        # Process each row
        for idx, row in df_raw.iterrows():
            # Skip empty rows
            if row.isna().all():
                continue

            name_term = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            lgd = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            
            # Skip header rows and PD row
            if name_term.lower() in ['pd', 'name/term', pd_rating.lower()]:
                continue
                
            # If we have a value in first column but no LGD, it's a parent
            if name_term and not lgd:
                if name_term.lower() != 'sub total':  # Skip subtotal rows
                    current_parent = name_term
                    rows.append({
                        'type': 'parent',
                        'parent': None,
                        'data': {
                            'Name/Term': name_term,
                            'LGD': '',
                            '% RR Used': None,
                            '% AGG Used': None,
                            'Used': None,
                            'Available': None,
                            'Total Exposure': None,
                            '% TE of RR': None,
                            '% TE of AGG': None
                        }
                    })
            # If we have LGD, it's a data row
            elif lgd:
                # Skip if this appears to be a header row
                if lgd.lower() == 'lgd':
                    continue
                    
                rows.append({
                    'type': 'data',
                    'parent': current_parent,
                    'data': {
                        'Name/Term': name_term,
                        'LGD': lgd,
                        '% RR Used': safe_numeric_convert(row.iloc[2]),
                        '% AGG Used': safe_numeric_convert(row.iloc[3]),
                        'Used': safe_numeric_convert(row.iloc[4]),
                        'Available': safe_numeric_convert(row.iloc[5]),
                        'Total Exposure': safe_numeric_convert(row.iloc[6]),
                        '% TE of RR': safe_numeric_convert(row.iloc[7]),
                        '% TE of AGG': safe_numeric_convert(row.iloc[8])
                    }
                })

        return pd_rating, [
            'Name/Term', 'LGD', '% RR Used', '% AGG Used', 'Used', 
            'Available', 'Total Exposure', '% TE of RR', '% TE of AGG'
        ], rows

    except Exception as e:
        raise Exception(f"Error processing data: {str(e)}")

def create_excel(category, headers, rows):
    """Create formatted Excel file"""
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            workbook = writer.book
            worksheet = workbook.create_sheet('Sheet1')
            
            # Write PD in cell A1
            worksheet.cell(row=1, column=1, value='PD')
            worksheet.cell(row=1, column=1).font = Font(bold=True)
            
            # Write headers in row 2
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=2, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
            # Write category in row 3
            cell = worksheet.cell(row=3, column=1, value=category)
            cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            
            # Styles
            yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Write data
            current_row = 4
            for row in rows:
                data = row['data']
                is_parent = row['type'] == 'parent'
                
                # Write each column
                for col, header in enumerate(headers, 1):
                    cell = worksheet.cell(row=current_row, column=col, value=data[header])
                    
                    # Parent row formatting
                    if is_parent:
                        cell.fill = yellow_fill
                        if col == 1:  # Only bold the Name/Term for parent rows
                            cell.font = Font(bold=True)
                    
                    # Formatting for numeric values
                    if header.startswith('%') and data[header] is not None:
                        cell.number_format = '0.00%'
                        cell.value = data[header] / 100 if data[header] else None
                    elif header in ['Used', 'Available', 'Total Exposure']:
                        cell.number_format = '#,##0'
                    
                    # Alignment
                    if header == 'LGD':
                        cell.alignment = Alignment(horizontal='center')
                    elif header in ['Used', 'Available', 'Total Exposure'] or header.startswith('%'):
                        cell.alignment = Alignment(horizontal='right')
                    
                    cell.border = thin_border
                
                current_row += 1
            
            # Set column widths
            worksheet.column_dimensions['A'].width = 40  # Name/Term
            worksheet.column_dimensions['B'].width = 10  # LGD
            for i in range(3, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(i)].width = 15

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
                        
                        # Process data
                        category, headers, rows = process_pd_data(df_raw)
                        
                        # Create Excel
                        excel_data = create_excel(category, headers, rows)
                        
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

