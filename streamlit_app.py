#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

def clean_and_structure_data(df):
    """Transform messy Excel data into clean structured format"""
    try:
        # Find the quality row
        quality_rows = df[df.iloc[:, 0].str.contains('Quality', na=False, case=False)]
        if quality_rows.empty:
            raise ValueError("Could not find Quality category row")
        
        category = df.iloc[quality_rows.index[0], 0]
        
        # Initialize structured data
        structured_data = []
        current_company = None
        
        # Skip header rows and start processing
        data_start = quality_rows.index[0] + 2  # Skip category and header row
        
        for idx in range(data_start, len(df)):
            row = df.iloc[idx]
            
            # Skip totally empty rows
            if row.isna().all():
                continue
                
            # Check if this is a company name (typically has LLC, INC, CORP, or no values in other columns)
            if (not pd.isna(row.iloc[0]) and 
                (pd.isna(row.iloc[1]) or any(x in str(row.iloc[0]).upper() for x in ['LLC', 'INC', 'CORP']))):
                current_company = row.iloc[0]
                continue
            
            # If we have data row with values
            if not pd.isna(row.iloc[1]) and isinstance(row.iloc[4], (int, float)):
                # Clean row data
                clean_row = {
                    'Name/Term': current_company,
                    'Detail': str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else '',
                    'LGD': row.iloc[1],
                    '% RR Used': pd.to_numeric(row.iloc[2], errors='coerce') / 100 if not pd.isna(row.iloc[2]) else 0,
                    '% AGG Used': pd.to_numeric(row.iloc[3], errors='coerce') / 100 if not pd.isna(row.iloc[3]) else 0,
                    'Used': pd.to_numeric(row.iloc[4], errors='coerce'),
                    'Available': pd.to_numeric(row.iloc[5], errors='coerce'),
                    'Total Exposure': pd.to_numeric(row.iloc[6], errors='coerce'),
                    '% TE of RR': pd.to_numeric(row.iloc[7], errors='coerce') / 100 if not pd.isna(row.iloc[7]) else 0,
                    '% TE of AGG': pd.to_numeric(row.iloc[8], errors='coerce') / 100 if not pd.isna(row.iloc[8]) else 0
                }
                structured_data.append(clean_row)
        
        # Create clean DataFrame
        clean_df = pd.DataFrame(structured_data)
        
        # Merge company name and detail
        clean_df['Name/Term'] = clean_df.apply(
            lambda x: x['Detail'] if pd.isna(x['Name/Term']) else 
                     f"{x['Name/Term']}\n{x['Detail']}" if x['Detail'] else 
                     x['Name/Term'], 
            axis=1
        )
        
        # Drop the detail column
        clean_df = clean_df.drop('Detail', axis=1)
        
        return category, clean_df
        
    except Exception as e:
        st.error(f"Error in data structuring: {str(e)}")
        raise

def create_styled_excel(category, df):
    """Create styled Excel file matching the first screenshot format"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
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
        
        # Apply formatting
        for row_idx, row in enumerate(df.itertuples(), start=3):
            # Format percentages
            for col in [4, 5, 9, 10]:  # Percentage columns
                cell = worksheet.cell(row=row_idx, column=col)
                cell.number_format = '0.00%'
            
            # Format numbers
            for col in [6, 7, 8]:  # Number columns
                cell = worksheet.cell(row=row_idx, column=col)
                cell.number_format = '#,##0'
            
            # Set row height for wrapped text
            worksheet.row_dimensions[row_idx].height = 30
        
        # Apply borders and column widths
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='center')
        
        # Set column widths
        worksheet.column_dimensions['A'].width = 40  # Name/Term
        worksheet.column_dimensions['B'].width = 10  # LGD
        for col in range(3, worksheet.max_column + 1):
            worksheet.column_dimensions[get_column_letter(col)].width = 15
    
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
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                    category, structured_df = clean_and_structure_data(df)
                    
                    # Show preview
                    st.subheader(f"Preview: {sheet_name}")
                    st.dataframe(structured_df)
                    
                    # Create download button
                    excel_data = create_styled_excel(category, structured_df)
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

