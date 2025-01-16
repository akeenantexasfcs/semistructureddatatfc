#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import json
import io
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

###############################################################################
# 1. Function to create a styled Excel from a specific JSON schema
###############################################################################
def create_styled_excel(processed_data):
    """
    Expects processed_data in the format:
    {
      "SheetName": {
        "category": "Some Category Title",
        "entries": [
          {
            "name": "Company A",
            "term": "Term1",
            "lgd": 0.25,
            "metrics": {
                "percentRRUsed": 0.1,
                "percentAGGUsed": 0.2,
                "used": 100,
                "available": 200,
                "totalExposure": 300,
                "percentTERR": 0.05,
                "percentTEAGG": 0.06
            },
            "sub_entries": [...]
          },
          ...
        ]
      },
      "AnotherSheet": { ... }
    }
    Creates a styled Excel with one sheet per top-level key.
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, data in processed_data.items():
            # Convert the data to rows
            rows = []
            for entry in data['entries']:
                # Add the company name row
                rows.append({
                    'Name/Term': entry['name'],
                    'LGD': '',
                    '% RR Used': '',
                    '% AGG Used': '',
                    'Used': '',
                    'Available': '',
                    'Total Exposure': '',
                    '% TE of RR': '',
                    '% TE of AGG': ''
                })
                
                # Add the metrics row
                if entry.get('metrics'):
                    rows.append({
                        'Name/Term': entry.get('term', ''),
                        'LGD': entry.get('lgd', ''),
                        '% RR Used': entry['metrics'].get('percentRRUsed', ''),
                        '% AGG Used': entry['metrics'].get('percentAGGUsed', ''),
                        'Used': entry['metrics'].get('used', ''),
                        'Available': entry['metrics'].get('available', ''),
                        'Total Exposure': entry['metrics'].get('totalExposure', ''),
                        '% TE of RR': entry['metrics'].get('percentTERR', ''),
                        '% TE of AGG': entry['metrics'].get('percentTEAGG', '')
                    })
                
                # Add sub-entries if any
                for sub in entry.get('sub_entries', []):
                    rows.append({
                        'Name/Term': sub.get('term', ''),
                        'LGD': sub.get('lgd', ''),
                        '% RR Used': sub['metrics'].get('percentRRUsed', ''),
                        '% AGG Used': sub['metrics'].get('percentAGGUsed', ''),
                        'Used': sub['metrics'].get('used', ''),
                        'Available': sub['metrics'].get('available', ''),
                        'Total Exposure': sub['metrics'].get('totalExposure', ''),
                        '% TE of RR': sub['metrics'].get('percentTERR', ''),
                        '% TE of AGG': sub['metrics'].get('percentTEAGG', '')
                    })
            
            # Create DataFrame and write to Excel
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Get the workbook and the worksheet
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Insert a title row
            worksheet.insert_rows(1)
            worksheet['A1'] = data['category']
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
            
            # Apply formatting to rows
            # Start after title row (1) and header row (2), so data starts at row 3
            for row_idx, row in enumerate(rows, start=3):
                # Highlight "company rows" (LGD is empty = 'company row')
                if row['LGD'] == '':
                    for col in range(1, worksheet.max_column + 1):
                        cell = worksheet.cell(row=row_idx, column=col)
                        cell.fill = yellow_fill
                        cell.font = Font(bold=True)
                else:
                    # Format percentages in columns 3, 4, 8, 9
                    for col in [3, 4, 8, 9]:
                        cell = worksheet.cell(row=row_idx, column=col)
                        cell.number_format = '0.00%'
                    # Format numbers in columns 5, 6, 7
                    for col in [5, 6, 7]:
                        cell = worksheet.cell(row=row_idx, column=col)
                        cell.number_format = '#,##0'
            
            # Apply borders to all cells in the used range
            for row in worksheet.iter_rows(min_row=1, max_row=len(rows) + 2):
                for cell in row:
                    cell.border = thin_border
            
            # Adjust column widths
            for col in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 30)

    output.seek(0)
    return output

###############################################################################
# 2. Main Streamlit app: Upload Excel -> Convert to JSON -> Display
#    Optionally, transform that JSON -> Download Styled Excel
###############################################################################
def main():
    st.title("Excel to JSON Converter and Styled Excel Generator")
    
    # 2.1 Upload an Excel file (xlsx or xls)
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx","xls"])
    
    if uploaded_file is not None:
        # 2.2 Read the Excel into a DataFrame
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            return
        
        # 2.3 Convert DataFrame to JSON (behind the scenes)
        data_json = df.to_json(orient='records')
        
        # 2.4 Display the DataFrame as a clean table
        st.subheader("Tabular Data from Excel")
        st.dataframe(df)
        
        # 2.5 Provide a download button to get the JSON
        st.download_button(
            label="Download JSON",
            data=data_json,
            file_name="excel_data.json",
            mime="application/json"
        )
        
        # 2.6 (Optional) If your Excel corresponds to the special JSON schema,
        #     parse the DataFrame into the required structure, then generate
        #     a styled Excel using `create_styled_excel`.
        #
        #     This parsing depends on how you want to map DataFrame -> processed_data.
        #     Below is just a minimal example for demonstration.
        
        # EXAMPLE: pretend we have a simple category plus 'entries' format
        processed_data_example = {
            "DemoSheet": {
                "category": "Demo Category Title",
                "entries": []
            }
        }
        
        # We'll just treat each row in df as an 'entry' with placeholders
        for idx, row in df.iterrows():
            processed_data_example["DemoSheet"]["entries"].append({
                "name": str(row[0]) if len(df.columns) > 0 else f"Row{idx}",
                "term": "",
                "lgd": "",
                "metrics": {
                    "percentRRUsed": 0.10,   # Hard-coded for demo
                    "percentAGGUsed": 0.20,
                    "used": 100,
                    "available": 200,
                    "totalExposure": 300,
                    "percentTERR": 0.05,
                    "percentTEAGG": 0.06
                },
                "sub_entries": []
            })
        
        # Generate a styled Excel using the example processed_data
        excel_data = create_styled_excel(processed_data_example)
        
        # 2.7 Provide a download button for the STYLED Excel
        st.download_button(
            label="Download Styled Excel (Example)",
            data=excel_data,
            file_name="styled_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.write("Please upload a file to proceed.")

if __name__ == "__main__":
    main()

