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
    
    #splitting date to seperate day, month and year
    from datetime import datetime

    # Try to extract and parse the Issue Opening Date
    issue_date_raw = data_dict.get("Issue Opening Date")
    if issue_date_raw:
        try:
            # Convert formats like 'April 9, 2025' to datetime object
            parsed_date = datetime.strptime(issue_date_raw, "%B %d, %Y")            
            data_dict['date_ddmmyyyy'] = parsed_date.strftime("%d%m%Y")
            #converting to individual d d m m y y y y cells
            data_dict['d1'] = data_dict['date_ddmmyyyy'][0]
            data_dict['d2'] = data_dict['date_ddmmyyyy'][1]
            data_dict['m1'] = data_dict['date_ddmmyyyy'][2]
            data_dict['m2'] = data_dict['date_ddmmyyyy'][3]
            data_dict['y1'] = data_dict['date_ddmmyyyy'][4]
            data_dict['y2'] = data_dict['date_ddmmyyyy'][5]
            data_dict['y3'] = data_dict['date_ddmmyyyy'][6]
            data_dict['y4'] = data_dict['date_ddmmyyyy'][7]
        
        
        except ValueError as e:
            print(f"⚠️ Could not parse Issue Opening Date: {issue_date_raw}")

    # Create DataFrame
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
# Enhanced function for better table structure handling
def convert_docx_to_xlsx(docx_file_path, xlsx_file_path, debug=True):
    from docx import Document
    import pandas as pd
    import re
    from datetime import datetime

    doc = Document(docx_file_path)
    
    data_dict = {}
    debug_info = []
    
    for table_idx, table in enumerate(doc.tables):
        if debug:
            print(f"\n📋 Processing Table {table_idx + 1}:")
        
        for row_idx, row in enumerate(table.rows):
            # Clean and extract all cell contents
            cells = []
            for cell in row.cells:
                cell_text = cell.text.strip().replace('\n', ' ').replace('\t', ' ')
                # Remove excessive whitespace
                cell_text = re.sub(r'\s+', ' ', cell_text).strip()
                cells.append(cell_text)
            
            # Skip completely empty rows
            if not any(cell.strip() for cell in cells):
                continue
            
            # Skip header rows (containing "TERMS OF ISSUE")
            if any("TERMS OF ISSUE" in cell.upper() for cell in cells):
                continue
            
            # Initialize extraction variables
            key = ""
            value = ""
            extraction_method = ""
            
            # Handle different column structures
            num_cols = len(cells)
            non_empty_cells = [cell for cell in cells if cell.strip()]
            
            if debug:
                print(f"  Row {row_idx + 1}: {num_cols} columns, {len(non_empty_cells)} non-empty")
                for i, cell in enumerate(cells):
                    if cell.strip():
                        print(f"    Col {i+1}: '{cell[:50]}{'...' if len(cell) > 50 else ''}'")
            
            # Strategy 1: Find key-value pairs in adjacent columns
            if num_cols >= 2:
                # Look for patterns where column i has key and column i+1 has value
                for i in range(num_cols - 1):
                    if cells[i].strip() and cells[i+1].strip():
                        # Skip if it looks like a continuation of previous content
                        if not any(skip_word in cells[i].lower() for skip_word in ['terms', 'issue', '>', '*']):
                            key = cells[i].strip()
                            value = cells[i+1].strip()
                            extraction_method = f"Col{i+1}→Col{i+2}"
                            break
            
            # Strategy 2: Handle single column with meaningful content
            if not key and num_cols >= 1:
                for i, cell in enumerate(cells):
                    if cell.strip() and not any(skip_word in cell.lower() for skip_word in ['terms', 'issue', '>', '*']):
                        # Check if this looks like a standalone value
                        if len(cell.strip()) > 3:  # Avoid very short strings
                            key = f"Field_{table_idx + 1}_{row_idx + 1}_{i + 1}"
                            value = cell.strip()
                            extraction_method = f"Single_Col{i+1}"
                            break
            
            # Strategy 3: Handle multi-column data (3+ columns)
            if not key and num_cols >= 3:
                # Look for key in first non-empty column, value in subsequent columns
                first_col_idx = None
                for i, cell in enumerate(cells):
                    if cell.strip():
                        first_col_idx = i
                        break
                
                if first_col_idx is not None and first_col_idx < num_cols - 1:
                    key = cells[first_col_idx].strip()
                    # Combine remaining non-empty columns as value
                    remaining_values = [cells[j].strip() for j in range(first_col_idx + 1, num_cols) if cells[j].strip()]
                    if remaining_values:
                        value = " | ".join(remaining_values)  # Use separator for multiple values
                        extraction_method = f"Multi_Col{first_col_idx+1}→{num_cols}"
            
            # Clean and validate the extracted key
            if key:
                # Remove problematic characters and normalize
                key = re.sub(r'[>\*\[\]{}]', '', key)  # Remove markup characters
                key = re.sub(r'\s+', ' ', key).strip()
                
                # Skip if key is too short or contains only common words
                if len(key) < 2 or key.lower() in ['', 'na', 'not applicable', 'terms', 'issue']:
                    key = ""
            
            # Store the data if we found a valid key-value pair
            if key and key.strip():
                clean_key = key.strip()
                clean_value = value.strip() if value else "Present"
                
                # Handle duplicate keys
                original_key = clean_key
                counter = 1
                while clean_key in data_dict:
                    clean_key = f"{original_key}_{counter}"
                    counter += 1
                
                data_dict[clean_key] = clean_value
                
                if debug:
                    debug_info.append({
                        'Table': table_idx + 1,
                        'Row': row_idx + 1,
                        'Method': extraction_method,
                        'Key': clean_key,
                        'Value': clean_value[:100] + '...' if len(clean_value) > 100 else clean_value,
                        'Original_Cells': ' | '.join([f"Col{i+1}: {cell}" for i, cell in enumerate(cells) if cell.strip()])
                    })
                    print(f"    ✓ Extracted: {clean_key} = {clean_value[:50]}{'...' if len(clean_value) > 50 else ''}")

    # Process special date fields
    date_fields_to_process = [
        "Issue Opening Date", "Issue Closing Date", "Pay-in-Date", 
        "Date of Allotment", "Deemed Date of Allotment", "Initial Fixing Date",
        "Final Fixing Date", "Redemption Date", "Coupon / Dividend payment dates"
    ]
    
    for date_field in date_fields_to_process:
        date_value = data_dict.get(date_field)
        if date_value:
            try:
                # Try different date formats
                date_formats = ["%B %d, %Y", "%B %d,%Y", "%d %B %Y", "%d/%m/%Y", "%Y-%m-%d"]
                parsed_date = None
                
                for fmt in date_formats:
                    try:
                        parsed_date = datetime.strptime(date_value, fmt)
                        break
                    except ValueError:
                        continue
                
                if parsed_date:
                    date_key = f"{date_field.replace(' ', '_').lower()}_formatted"
                    data_dict[date_key] = parsed_date.strftime("%d%m%Y")
                    
                    # For Issue Opening Date, create individual digit fields
                    if date_field == "Issue Opening Date":
                        formatted_date = parsed_date.strftime("%d%m%Y")
                        data_dict['d1'] = formatted_date[0]
                        data_dict['d2'] = formatted_date[1]
                        data_dict['m1'] = formatted_date[2]
                        data_dict['m2'] = formatted_date[3]
                        data_dict['y1'] = formatted_date[4]
                        data_dict['y2'] = formatted_date[5]
                        data_dict['y3'] = formatted_date[6]
                        data_dict['y4'] = formatted_date[7]
                        
                        if debug:
                            print(f"📅 Processed {date_field}: {date_value} → {formatted_date}")
                
            except Exception as e:
                if debug:
                    print(f"⚠️ Could not parse date field {date_field}: {date_value} - {e}")

    # Create DataFrames
    if data_dict:
        df = pd.DataFrame([data_dict])
        if debug:
            print(f"\n📊 Created DataFrame with {len(data_dict)} columns")
    else:
        df = pd.DataFrame()
        if debug:
            print("⚠️ No data extracted - empty DataFrame created")
    
    # Write to Excel with enhanced formatting
    with pd.ExcelWriter(xlsx_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Terms", index=False)
        
        # Add debug sheet if requested
        if debug and debug_info:
            debug_df = pd.DataFrame(debug_info)
            debug_df.to_excel(writer, sheet_name="Debug_Log", index=False)
        
        # Enhanced formatting
        workbook = writer.book
        
        if not df.empty:
            worksheet = writer.sheets['Terms']
            
            # Header formatting
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#D7E4BC',
                'border': 1
            })
            
            # Data formatting
            data_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'top',
                'border': 1
            })
            
            # Apply formatting and adjust column widths
            for col_num, column_name in enumerate(df.columns.values):
                worksheet.write(0, col_num, column_name, header_format)
                
                # Calculate appropriate column width
                max_length = max(
                    len(str(column_name)),
                    df[column_name].astype(str).map(len).max() if not df.empty else 0
                )
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                worksheet.set_column(col_num, col_num, adjusted_width, data_format)
        
        # Format debug sheet if it exists
        if debug and debug_info:
            debug_worksheet = writer.sheets['Debug_Log']
            debug_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#FFE6CC',
                'border': 1
            })
            
            for col_num, column_name in enumerate(debug_df.columns.values):
                debug_worksheet.write(0, col_num, column_name, debug_header_format)
                max_length = max(
                    len(str(column_name)),
                    debug_df[column_name].astype(str).map(len).max() if not debug_df.empty else 0
                )
                debug_worksheet.set_column(col_num, col_num, min(max_length + 2, 60))

    print(f"\n✅ Export complete!")
    print(f"📊 Total fields extracted: {len(data_dict)}")
    print(f"📁 Output file: {xlsx_file_path}")
    
    if debug:
        print(f"🔍 Debug log with {len(debug_info)} entries saved to 'Debug_Log' sheet")
        print("\n📋 Sample extracted data:")
        for i, (key, value) in enumerate(list(data_dict.items())[:10]):
            print(f"  {i+1}. {key}: {value[:100]}{'...' if len(str(value)) > 100 else ''}")
    
    return data_dict

# Enhanced usage function
# def process_document_with_validation(docx_path, xlsx_path, debug=True):
#     """
#     Process document with additional validation and error handling
#     """
#     try:
#         print(f"🔄 Starting conversion: {docx_path} → {xlsx_path}")
#         result = convert_docx_to_xlsx(docx_path, xlsx_path, debug=debug)
        
#         if result:
#             print(f"✅ Successfully extracted {len(result)} fields")
#             return result
#         else:
#             print("⚠️ No data was extracted from the document")
#             return {}
            
    # except Exception as e:
    #     print(f"❌ Error processing document: {str(e)}")
    #     return {}

# Usage examples:
"""
# Basic usage
data = convert_docx_to_xlsx('Term Sheet - Series 129 for claude.docx', 'output.xlsx')

# With comprehensive debugging
data = convert_docx_to_xlsx('Term Sheet - Series 129 for claude.docx', 'output_debug.xlsx', debug=True)

# Using the validation wrapper
data = process_document_with_validation('Term Sheet - Series 129 for claude.docx', 'validated_output.xlsx')

# Print extracted data summary
if data:
    print("\\nExtracted Fields:")
    for key, value in data.items():
        print(f"{key}: {value}")
"""


def save_uploaded_file(file, upload_folder):
    import os
    file_path = os.path.join(upload_folder, file.filename)
    file.save(file_path)
    return file_path