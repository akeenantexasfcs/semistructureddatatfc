#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def process_excel_sheet(df):
    """Process a single PD sheet from the Excel workbook."""
    try:
        # Find the quality row (contains 'Quality' in first column)
        quality_rows = df[df.iloc[:, 0].str.contains('Quality', na=False, case=False)]
        if quality_rows.empty:
            raise ValueError("Could not find Quality category row")
        
        category_row = quality_rows.index[0]
        category = df.iloc[category_row, 0]
        
        # Get headers
        headers = [
            'Name/Term', 'LGD', '% Used of RR', '% Used of AGG',
            'Used', 'Available', 'Total Exposure', '% TE of RR', '% TE of AGG'
        ]
        
        # Process data rows
        processed_rows = []
        current_company = None
        
        for idx in range(category_row + 1, len(df)):
            row = df.iloc[idx]
            
            # Skip empty rows
            if pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]):
                continue
                
            # Check if this is a company name row
            if not pd.isna(row.iloc[0]) and (pd.isna(row.iloc[1]) or row.iloc[1] == ''):
                current_company = row.iloc[0]
                new_row = [current_company] + [''] * (len(headers) - 1)
                processed_rows.append(new_row)
                continue
            
            # Process data rows
            if not pd.isna(row.iloc[1]):  # Has LGD value
                data_row = [
                    row.iloc[0] if not pd.isna(row.iloc[0]) else '',  # Name/Term
                    row.iloc[1],  # LGD
                    row.iloc[2] if not pd.isna(row.iloc[2]) else 0,  # % Used of RR
                    row.iloc[3] if not pd.isna(row.iloc[3]) else 0,  # % Used of AGG
                    row.iloc[4] if not pd.isna(row.iloc[4]) else 0,  # Used
                    row.iloc[5] if not pd.isna(row.iloc[5]) else 0,  # Available
                    row.iloc[6] if not pd.isna(row.iloc[6]) else 0,  # Total Exposure
                    row.iloc[7] if not pd.isna(row.iloc[7]) else 0,  # % TE of RR
                    row.iloc[8] if not pd.isna(row.iloc[8]) else 0,  # % TE of AGG
                ]
                processed_rows.append(data_row)
        
        # Create DataFrame
        processed_df = pd.DataFrame(processed_rows, columns=headers)
        return category, processed_df
        
    except Exception as e:
        st.error(f"Error in process_excel_sheet: {str(e)}")
        raise

def create_styled_excel(category, df):
    """Create styled Excel file"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write DataFrame
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Add category row
        worksheet.insert_rows(1)
        worksheet['A1'] = category
        worksheet['A1'].font = Font(bold=True)
        
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
        
        # Format rows
        for row_idx, row in enumerate(df.itertuples(), start=3):
            # Company row (no LGD)
            if pd.isna(row.LGD) or row.LGD == '':
                for col in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_idx, column=col)
                    cell.fill = yellow_fill
                    cell.font = Font(bold=True)
            else:
                # Format percentages
                for col in [3, 4, 8, 9]:  # Percentage columns
                    cell = worksheet.cell(row=row_idx, column=col)
                    if cell.value:
                        cell.number_format = '0.00%'
                
                # Format numbers
                for col in [5, 6, 7]:  # Number columns
                    cell = worksheet.cell(row=row_idx, column=col)
                    if cell.value:
                        cell.number_format = '#,##0'
        
        # Borders and column widths
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
                    # Read with header=None to get raw data
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
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

