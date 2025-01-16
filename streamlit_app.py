#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import json
import io
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def process_excel_sheet(df):
    """Process a single PD sheet from the Excel workbook."""
    # Find the category row (usually contains 'Quality')
    category_row = df[df.iloc[:, 0].str.contains('Quality', na=False)].index[0]
    category = df.iloc[category_row, 0]
    
    # Find the header row (usually contains 'LGD')
    header_row = df[df.iloc[:, 1].str.contains('LGD', na=False)].index[0]
    
    # Set the headers
    headers = df.iloc[header_row]
    df.columns = headers
    
    # Get data after headers
    data_df = df.iloc[header_row + 1:].copy()
    data_df.columns = headers
    
    # Clean up column names
    data_df.columns = data_df.columns.str.strip()
    
    return category, data_df

def create_styled_excel(category, df):
    """Create styled Excel file similar to the screenshot"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write the DataFrame
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Add title row
        worksheet.insert_rows(1)
        worksheet['A1'] = category
        worksheet['A1'].font = Font(bold=True)
        
        # Define styles
        yellow_fill = PatternFill(start_color='FFEB9C',
                                end_color='FFEB9C',
                                fill_type='solid')
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Apply formatting
        for row_idx, row in enumerate(df.itertuples(), start=3):
            # Check if this is a main entry row (no LGD value)
            if pd.isna(row.LGD):
                for col in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_idx, column=col)
                    cell.fill = yellow_fill
                    cell.font = Font(bold=True)
            else:
                # Format percentages
                percent_cols = ['% RR Used', '% AGG Used', '% TE of RR']
                for col_name in percent_cols:
                    col_idx = df.columns.get_loc(col_name) + 1
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.number_format = '0.00%'
                
                # Format numbers
                number_cols = ['Used', 'Available', 'Total Exposure']
                for col_name in number_cols:
                    col_idx = df.columns.get_loc(col_name) + 1
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.number_format = '#,##0'
        
        # Apply borders and adjust column widths
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
        
        for col in worksheet.columns:
            max_length = 0
            column = get_column_letter(col[0].column)
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[column].width = min(max_length + 2, 30)
    
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
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                    category, processed_df = process_excel_sheet(df)
                    
                    # Show preview
                    st.subheader(f"Preview: {sheet_name}")
                    st.dataframe(processed_df)
                    
                    # Create download button
                    excel_data = create_styled_excel(category, processed_df)
                    st.download_button(
                        label=f"Download {sheet_name} Excel",
                        data=excel_data,
                        file_name=f"{sheet_name}_formatted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.write("Please make sure the Excel file follows the expected format.")

if __name__ == "__main__":
    main()

