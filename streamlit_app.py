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
        # Remove commas and convert to float
        if isinstance(value, str):
            value = value.replace(',', '')
        return float(value)
    except (ValueError, TypeError):
        return None

def process_pd_data(df_raw):
    """Process the PD sheet using fixed positions instead of searching"""
    try:
        # Check if DataFrame has enough rows
        if df_raw.shape[0] < 2:
            raise ValueError("Excel sheet does not have enough rows. Expected PD in cell A2.")
            
        # 1. Get category from cell A2 (index 1,0)
        try:
            category = str(df_raw.iloc[1, 0]).strip()
            if pd.isna(category) or category == "":
                raise ValueError("No PD found in cell A2")
        
        # 2. Get headers from row 2 (index 2)
        headers = [
            'Name/Term', 'LGD', '% Used of RR', '% Used of AGG',
            'Used', 'Available', 'Total Exposure', '% TE of RR', '% TE of AGG'
        ]
        
        # Check if we have enough rows for data
        if df_raw.shape[0] < 4:  # Need at least 4 rows (0-based index: 0,1,2,3)
            raise ValueError("Excel sheet does not have enough rows for data processing")
            
        # 3. Create DataFrame from row 3 onwards
        df_data = df_raw.iloc[3:].copy()
        df_data.columns = headers
        
        # 4. Drop completely empty rows
        df_data.dropna(how='all', inplace=True)
        
        # 5. Process numeric columns
        numeric_cols = ['% Used of RR', '% Used of AGG', 'Used', 
                       'Available', 'Total Exposure', '% TE of RR', '% TE of AGG']
        for col in numeric_cols:
            df_data[col] = df_data[col].apply(safe_numeric_convert)
        
        # 6. Clean text columns
        df_data['Name/Term'] = df_data['Name/Term'].fillna('').astype(str).str.strip()
        df_data['LGD'] = df_data['LGD'].fillna('').astype(str).str.strip()
        
        # 7. Identify parent rows (rows where Name/Term is filled but LGD is empty)
        parent_rows = df_data['Name/Term'].notna() & (df_data['LGD'].str.len() == 0)
        
        # 8. Create structured data for Excel formatting
        structured_data = []
        for idx, row in df_data.iterrows():
            structured_data.append({
                'row_type': 'parent' if parent_rows.iloc[idx] else 'data',
                'data': row.to_dict()
            })
        
        return category, headers, structured_data
    
    except Exception as e:
        raise Exception(f"Error processing data: {str(e)}")

def create_excel(category, headers, structured_data):
    """Create formatted Excel file with consistent styling"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        worksheet = workbook.create_sheet('Sheet1')
        
        # Write category in A2
        worksheet.cell(row=2, column=1, value=category)
        worksheet.cell(row=2, column=1).font = Font(bold=True)
        
        # Write headers in row 3
        for col, header in enumerate(headers, 1):
            cell = worksheet.cell(row=3, column=col, value=header)
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
        
        # Write data starting from row 4
        current_row = 4
        for item in structured_data:
            row_data = item['data']
            
            for col, header in enumerate(headers, 1):
                cell = worksheet.cell(row=current_row, column=col, value=row_data[header])
                
                # Apply styling
                if item['row_type'] == 'parent':
                    cell.fill = yellow_fill
                    cell.font = Font(bold=True)
                
                # Format numbers and percentages
                if header.startswith('%'):
                    if isinstance(row_data[header], (int, float)):
                        cell.number_format = '0.00%'
                        cell.value = row_data[header] / 100 if row_data[header] is not None else None
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
        worksheet.column_dimensions['A'].width = 40  # Name/Term
        worksheet.column_dimensions['B'].width = 10  # LGD
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
                    # Read sheet without header
                    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                    
                    # Process data using fixed positions
                    category, headers, structured_data = process_pd_data(df_raw)
                    
                    # Create formatted Excel
                    excel_data = create_excel(category, headers, structured_data)
                    
                    # Provide download button
                    st.download_button(
                        label=f"Download {sheet_name}",
                        data=excel_data,
                        file_name=f"{sheet_name}_formatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Show preview
                    preview_df = pd.DataFrame([item['data'] for item in structured_data])
                    st.dataframe(preview_df)
                    
        except Exception as e:
            error_msg = str(e)
            st.error(f"Error: {error_msg}")
            
            # Provide more specific guidance based on the error
            if "not have enough rows" in error_msg:
                st.write("The Excel sheet appears to be empty or doesn't have enough rows. Please ensure:")
                st.write("1. The PD is in cell A2")
                st.write("2. Headers are in row 3")
                st.write("3. Data starts from row 4")
            elif "No PD found in cell A2" in error_msg:
                st.write("Could not find PD information in cell A2. Please check the sheet format.")
            else:
                st.write("Please ensure the Excel file follows the expected format with PD in cell A2.")

if __name__ == "__main__":
    main()

