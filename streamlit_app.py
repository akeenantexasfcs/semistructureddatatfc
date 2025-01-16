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
    try:
        if pd.isna(value):
            return None
        float_val = float(str(value).replace(',', ''))
        return float_val
    except (ValueError, TypeError):
        return None

def process_pd_data(df):
    """Process the PD sheet and return structured data"""
    try:
        # Find the quality category
        category = None
        for idx, row in df.iterrows():
            if 'Quality' in str(row.iloc[0]):
                category = row.iloc[0]
                break

        # Get the header row index (where "LGD" appears)
        header_idx = None
        for idx, row in df.iterrows():
            if 'LGD' in str(row.iloc[1]):
                header_idx = idx
                break

        if header_idx is None:
            raise ValueError("Could not find header row with 'LGD'")

        # Set up headers
        headers = ['Name/Term', 'LGD', '% Used of RR', '% Used of AGG', 
                  'Used', 'Available', 'Total Exposure', '% TE of RR', '% TE of AGG']
        
        # Process data rows
        processed_data = []
        data_rows = df.iloc[header_idx+1:].values
        
        for row in data_rows:
            # Skip completely empty rows
            if all(pd.isna(cell) for cell in row):
                continue
                
            # Check if this is a title/parent row
            if pd.notna(row[0]) and (pd.isna(row[1]) or str(row[1]).strip() == ''):
                processed_data.append({
                    'row_type': 'parent',
                    'data': {
                        'Name/Term': str(row[0]).strip(),
                        'LGD': '',
                        '% Used of RR': '',
                        '% Used of AGG': '',
                        'Used': '',
                        'Available': '',
                        'Total Exposure': '',
                        '% TE of RR': '',
                        '% TE of AGG': ''
                    }
                })
            # Handle regular data rows
            elif pd.notna(row[1]):
                processed_data.append({
                    'row_type': 'data',
                    'data': {
                        'Name/Term': str(row[0]).strip() if pd.notna(row[0]) else '',
                        'LGD': str(row[1]).strip(),
                        '% Used of RR': safe_numeric_convert(row[2]),
                        '% Used of AGG': safe_numeric_convert(row[3]),
                        'Used': safe_numeric_convert(row[4]),
                        'Available': safe_numeric_convert(row[5]),
                        'Total Exposure': safe_numeric_convert(row[6]),
                        '% TE of RR': safe_numeric_convert(row[7]),
                        '% TE of AGG': safe_numeric_convert(row[8])
                    }
                })
        
        return category, headers, processed_data
    except Exception as e:
        raise Exception(f"Error processing data: {str(e)}")

def create_excel(category, headers, processed_data):
    """Create formatted Excel file"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Create workbook and worksheet
        workbook = writer.book
        worksheet = workbook.create_sheet('Sheet1')
        
        # Write category
        worksheet.cell(row=1, column=1, value=category)
        worksheet.cell(row=1, column=1).font = Font(bold=True)
        
        # Write headers
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
        for item in processed_data:
            row_data = item['data']
            
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=current_row, column=col, value=row_data[header])
                
                # Apply styling
                if item['row_type'] == 'parent':
                    cell.fill = yellow_fill
                    cell.font = Font(bold=True)
                else:
                    # Format numbers and percentages
                    if header in ['% Used of RR', '% Used of AGG', '% TE of RR', '% TE of AGG']:
                        if isinstance(row_data[header], (int, float)):
                            cell.number_format = '0.00%'
                            cell.value = row_data[header] / 100
                    elif header in ['Used', 'Available', 'Total Exposure']:
                        if isinstance(row_data[header], (int, float)):
                            cell.number_format = '#,##0'
                
                # Alignment
                if header == 'Name/Term':
                    cell.alignment = Alignment(horizontal='left')
                elif header == 'LGD':
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.alignment = Alignment(horizontal='right')
                
                cell.border = thin_border
            
            current_row += 1
        
        # Set column widths
        worksheet.column_dimensions['A'].width = 40
        worksheet.column_dimensions['B'].width = 10
        for i in range(3, len(headers) + 1):
            worksheet.column_dimensions[get_column_letter(i)].width = 15

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
                    # Read sheet without header
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                    
                    # Process data
                    category, headers, processed_data = process_pd_data(df)
                    
                    # Create Excel file
                    excel_data = create_excel(category, headers, processed_data)
                    
                    # Download button
                    st.download_button(
                        label=f"Download {sheet_name}",
                        data=excel_data,
                        file_name=f"{sheet_name}_formatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Show preview
                    preview_df = pd.DataFrame([item['data'] for item in processed_data])
                    st.dataframe(preview_df)
                    
        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.write("Please make sure the Excel file follows the expected format.")

if __name__ == "__main__":
    main()

