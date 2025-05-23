def convert_docx_to_xlsx_improved(docx_file_path, xlsx_file_path):
    from docx import Document
    import pandas as pd
    import re

    doc = Document(docx_file_path)
    
    # Collect all key-value pairs from all tables
    data_dict = {}
    
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ').replace('\t', ' ') for cell in row.cells]
            
            # Skip completely empty rows
            if not any(cell.strip() for cell in cells):
                continue
            
            # Skip header rows (like "TERMS OF ISSUE")
            if len(cells) >= 2 and cells[1] and cells[1].upper() == "TERMS OF ISSUE":
                continue
            
            # Process rows with at least 2 columns
            if len(cells) >= 2:
                key = ""
                value = ""
                
                # Case 1: Key in second column, value in third column
                if cells[1].strip() and len(cells) > 2:
                    key = cells[1].strip()
                    value = cells[2].strip() if cells[2].strip() else "Not specified"
                
                # Case 2: Key in first column, value in second column
                elif cells[0].strip() and cells[1].strip():
                    key = cells[0].strip()
                    value = cells[1].strip()
                
                # Case 3: Only second column has content (use as both key and value)
                elif cells[1].strip() and not cells[0].strip() and (len(cells) <= 2 or not cells[2].strip()):
                    key = cells[1].strip()
                    value = "Yes" if key else ""
                
                # Clean up the key and add to dictionary if valid
                if key and key not in ["", "TERMS OF ISSUE"]:
                    # Clean up key name - remove extra spaces and special characters
                    clean_key = re.sub(r'\s+', ' ', key).strip()
                    
                    # Handle duplicate keys by adding suffix
                    original_key = clean_key
                    counter = 1
                    while clean_key in data_dict:
                        clean_key = f"{original_key}_{counter}"
                        counter += 1
                    
                    data_dict[clean_key] = value
    
    # Convert dictionary to DataFrame (transpose to make keys as columns)
    if data_dict:
        df = pd.DataFrame([data_dict])
    else:
        df = pd.DataFrame()
    
    # Write to Excel with formatting
    with pd.ExcelWriter(xlsx_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Terms", index=False)
        
        # Get workbook and worksheet objects for formatting
        workbook = writer.book
        worksheet = writer.sheets['Terms']
        
        # Add formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#D7E4BC'
        })
        
        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top'
        })
        
        # Apply header formatting
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Apply cell formatting and adjust column widths
        for col_num, column in enumerate(df.columns):
            max_length = max(
                df[column].astype(str).map(len).max(),  # Max length in column
                len(str(column))  # Length of column name
            )
            # Set column width (with some padding)
            worksheet.set_column(col_num, col_num, min(max_length + 2, 50))
        
        # Set row height for better readability
        worksheet.set_row(0, 30)  # Header row
        for row_num in range(1, len(df) + 1):
            worksheet.set_row(row_num, 20)
    writer.close()
    print(f"✅ Export complete!")
    print(f"📊 Total fields extracted: {len(data_dict)}")
    print(f"📄 Saved to: {xlsx_file_path}")
    
    return data_dict

# Alternative function for better debugging and validation
def convert_docx_to_xlsx(docx_file_path, xlsx_file_path, debug=True):
    from docx import Document
    import pandas as pd
    import re

    doc = Document(docx_file_path)
    
    data_dict = {}
    debug_info = []
    
    for table_idx, table in enumerate(doc.tables):
        if debug:
            print(f"\n📋 Processing Table {table_idx + 1}:")
        
        for row_idx, row in enumerate(table.rows):
            cells = [cell.text.strip().replace('\n', ' ').replace('\t', ' ') for cell in row.cells]
            
            # Skip empty rows
            if not any(cell.strip() for cell in cells):
                continue
            
            # Skip header rows
            if len(cells) >= 2 and cells[1] and cells[1].upper() == "TERMS OF ISSUE":
                continue
            
            if len(cells) >= 2:
                key = ""
                value = ""
                extraction_method = ""
                
                # Method 1: Key in col 2, value in col 3
                if cells[1].strip() and len(cells) > 2 and cells[2].strip():
                    key = cells[1].strip()
                    value = cells[2].strip()
                    extraction_method = "Col2→Col3"
                
                # Method 2: Key in col 1, value in col 2
                elif cells[0].strip() and cells[1].strip():
                    key = cells[0].strip()
                    value = cells[1].strip()
                    extraction_method = "Col1→Col2"
                
                # Method 3: Only col 2 has content
                elif cells[1].strip() and not cells[0].strip():
                    key = cells[1].strip()
                    value = "Present"
                    extraction_method = "Col2 only"
                
                if key and key not in ["", "TERMS OF ISSUE"]:
                    clean_key = re.sub(r'\s+', ' ', key).strip()
                    
                    # Handle duplicates
                    original_key = clean_key
                    counter = 1
                    while clean_key in data_dict:
                        clean_key = f"{original_key}_{counter}"
                        counter += 1
                    
                    data_dict[clean_key] = value
                    
                    if debug:
                        debug_info.append({
                            'Table': table_idx + 1,
                            'Row': row_idx + 1,
                            'Method': extraction_method,
                            'Key': clean_key,
                            'Value': value[:50] + '...' if len(value) > 50 else value
                        })
    
    #splitting date to seperate day, month and year
    from datetime import datetime

    # Try to extract and parse the Issue Opening Date
    issue_date_raw = data_dict.get("Issue Opening Date")
    if issue_date_raw:
        try:
            # Convert formats like 'April 9, 2025' to datetime object
            parsed_date = datetime.strptime(issue_date_raw, "%B %d, %Y")
            
            # Add individual components
            data_dict['date_day'] = f"{parsed_date.day:02d}"
            data_dict['date_month'] = f"{parsed_date.month:02d}"
            data_dict['date_year'] = f"{parsed_date.year:04d}"
            
            # Optional: Also add full compact version if needed (for checking)
            data_dict['date_ddmmyyyy'] = parsed_date.strftime("%d%m%Y")
        
        except ValueError as e:
            print(f"⚠️ Could not parse Issue Opening Date: {issue_date_raw}")

    # Create DataFrame
    if data_dict:
        df = pd.DataFrame([data_dict])
        
        # Also create debug DataFrame if requested
        if debug and debug_info:
            debug_df = pd.DataFrame(debug_info)
    else:
        df = pd.DataFrame()
    
    # Write to Excel
    with pd.ExcelWriter(xlsx_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Terms", index=False)
        
        if debug and debug_info:
            debug_df.to_excel(writer, sheet_name="Debug_Log", index=False)
        
        # Formatting (same as above)
        workbook = writer.book
        worksheet = writer.sheets['Terms']
        
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#D7E4BC'
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            max_length = max(
                df[df.columns[col_num]].astype(str).map(len).max() if not df.empty else 0,
                len(str(value))
            )
            worksheet.set_column(col_num, col_num, min(max_length + 2, 50))
    writer.close()
    print(f"✅ Export complete!")
    print(f"📊 Total fields extracted: {len(data_dict)}")
    if debug:
        print(f"🔍 Debug log with {len(debug_info)} entries saved to 'Debug_Log' sheet")
    
    return data_dict

# Usage examples:
"""
# Basic usage
data = convert_docx_to_xlsx_improved('your_document.docx', 'output.xlsx')

# With debugging to see what's being extracted
data = convert_docx_to_xlsx_with_debug('your_document.docx', 'output_with_debug.xlsx', debug=True)

# Print some extracted data
for key, value in list(data.items())[:10]:  # First 10 items
    print(f"{key}: {value}")
"""


def save_uploaded_file(file, upload_folder):
    import os
    file_path = os.path.join(upload_folder, file.filename)
    file.save(file_path)
    return file_path