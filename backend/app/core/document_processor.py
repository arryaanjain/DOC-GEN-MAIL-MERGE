from docx import Document
import pandas as pd
import re
from datetime import datetime
from app.utils.text_utils import extract_series_number, extract_issue_size_number
from app.utils.date_utils import handle_processing_date
from app.utils.number_utils import extract_tenor_days_number, extract_face_value_number, calculate_amount_raised, extract_discount_value_number, extract_issue_price_number
"""
OLD CODE: WILL BE MARKED TO DELETE LATER
"""
class DocumentProcessor:
    def __init__(self, debug=True):
        self.debug = debug

    # Enhanced function for better table structure handling
    def convert_docx_to_xlsx(self, docx_file_path, xlsx_file_path, processing_date=None, debug=True):

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
                    for i in range(num_cols - 1):
                        if cells[i].strip():
                            # Only skip exact matches of these phrases
                            skip_phrases = [
                                'terms of issue',
                                'terms and conditions',
                                '>>',
                                '***',
                                '*****'
                            ]
                            # Don't skip if it's a valid header
                            valid_headers = [
                                'issue price',
                                'issue opening date',
                                'issue closing date',
                                'discount at which security is issued',
                                'coupon'
                            ]
                            cell_lower = cells[i].lower()
                            
                            if (not any(phrase in cell_lower for phrase in skip_phrases) or 
                                any(header in cell_lower for header in valid_headers)):
                                key = cells[i].strip()
                                
                                # Combine values from column 2 and 3 if they're different
                                if num_cols > 2 and cells[1].strip() and cells[2].strip():
                                    if cells[1].strip() != cells[2].strip():
                                        value = f"{cells[1].strip()} | {cells[2].strip()}"
                                    else:
                                        value = cells[1].strip()
                                else:
                                    value = cells[i+1].strip()
                                    
                                extraction_method = f"Col{i+1}→Col{i+2}"

                                 # Add to data_dict right after extraction
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
                                        print(f"    ✓ Strategy 1: Added {clean_key} = {clean_value[:50]}...")
                                        debug_info.append({
                                            'Table': table_idx + 1,
                                            'Row': row_idx + 1,
                                            'Method': extraction_method,
                                            'Key': clean_key,
                                            'Value': clean_value[:100] + '...' if len(clean_value) > 100 else clean_value,
                                            'Original_Cells': ' | '.join([f"Col{i+1}: {cell}" for i, cell in enumerate(cells) if cell.strip()])
                                        })
                                break
                    
                # Strategy 2: Handle single column with meaningful content
                    if not key and num_cols >= 1:
                        skip_phrases = [
                            'terms of issue',
                            'terms and conditions',
                            '***',
                            '>>>'
                        ]
                        valid_headers = [
                            'issue price',
                            'issue opening date',
                            'issue closing date',
                            'discount at which security is issued'
                        ]
                        
                        for i, cell in enumerate(cells):
                            cell_lower = cell.strip().lower()
                            if (cell.strip() and 
                                not any(skip in cell_lower for skip in skip_phrases) or
                                any(header in cell_lower for header in valid_headers)):
                                if len(cell.strip()) > 3:  # Avoid very short strings
                                    key = cell.strip()  # Use actual content instead of Field_X_Y_Z
                                    value = cell.strip()
                                    extraction_method = f"Single_Col{i+1}"

                                    # Add to data_dict right after extraction
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
                                            # Add debug info
                                            debug_info.append({
                                                'Table': table_idx + 1,
                                                'Row': row_idx + 1,
                                                'Method': extraction_method,
                                                'Key': clean_key,
                                                'Value': clean_value[:100] + '...' if len(clean_value) > 100 else clean_value,
                                                'Original_Cells': ' | '.join([f"Col{i+1}: {cell}" for i, cell in enumerate(cells) if cell.strip()])
                                            })
                                            print(f"    ✓ Strategy 2: Added {clean_key} = {clean_value[:50]}...")
                                    break

                # Strategy 3: Handle multi-column data (3+ columns)
                if not key and num_cols >= 3:
                    if debug:
                        print("\n→ Entering Strategy 3")
                        print(f"    Number of columns: {num_cols}")
                        print(f"    Cells content: {cells}")

                    # Check if this is a Coupon row
                    first_cell = cells[0].strip().lower() if cells[0] else ""
                    
                    if first_cell == 'coupon':
                        if debug:
                            print("    ✓ Found Coupon row!")
                        
                        # Initialize coupon data structure if not exists
                        if 'coupon_data' not in data_dict:
                            data_dict['coupon_data'] = []
                        
                        # Get description and value from columns 2 and 3
                        description = cells[1].strip() if len(cells) > 1 else ""
                        value = cells[2].strip() if len(cells) > 2 else ""
                        
                        if description or value:
                            # Store complete information
                            combined_value = ""
                            if description and value:
                                combined_value = f"{description} | {value}"
                            elif description:
                                combined_value = description
                            elif value:
                                combined_value = value
                            
                            if combined_value:
                                data_dict['coupon_data'].append(combined_value)
                                
                                # Determine if this is first or second coupon entry
                                coupon_key = 'Coupon' if 'Coupon' not in data_dict else 'Coupon_1'
                                data_dict[coupon_key] = combined_value
                                
                                if debug:
                                    # Add debug info for coupon
                                    debug_info.append({
                                        'Table': table_idx + 1,
                                        'Row': row_idx + 1,
                                        'Method': 'Strategy3_Coupon',
                                        'Key': coupon_key,
                                        'Value': combined_value[:100] + '...' if len(combined_value) > 100 else combined_value,
                                        'Original_Cells': ' | '.join([f"Col{i+1}: {cell}" for i, cell in enumerate(cells) if cell.strip()])
                                    })
                                    print(f"    ✓ Added {coupon_key}: {combined_value[:50]}...")
                        
                        # Set method for debug info
                        extraction_method = "Coupon_Special"
                        key = None  # Prevent further processing of this row
                        
                        continue
                                            
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
                        if processing_date:
                            date_components = handle_processing_date(processing_date)
                            if date_components:
                                data_dict.update(date_components)
                                if debug:
                                    print(f"📅 Added processing date: {processing_date}")

                                
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
        # Extract series number from Product Code
        if 'Product Code' in data_dict:
            series_number = extract_series_number(data_dict['Product Code'])
            data_dict['series_number'] = series_number
            if debug:
                print(f"📎 Extracted series number: {series_number}")
        # Extract issue size number from Issue Size
        if 'Issue Size' in data_dict:
            issue_size_num, formatted_amount = extract_issue_size_number(data_dict['Issue Size'])
            data_dict['issue_size_number'] = issue_size_num
            data_dict['total_amount'] = formatted_amount
            if debug:
                print(f"💰 Extracted issue size number: {issue_size_num}")
                print(f"💰 Extracted total amount: {formatted_amount}")
        # Extract tenor days number
        if 'Tenor In Days' in data_dict:
            tenor_days = extract_tenor_days_number(data_dict['Tenor In Days'])
            data_dict['tenor_days_num'] = tenor_days
            if debug:
                print(f"📅 Extracted tenor days: {tenor_days}")
        # Extract face value number
        if 'Face Value' in data_dict:
            face_value_num, formatted_face_value = extract_face_value_number(data_dict['Face Value'])
            data_dict['face_value_num'] = face_value_num
            data_dict['formatted_face_value'] = formatted_face_value
            if debug:
                print(f"💵 Extracted face value: {face_value_num} ({formatted_face_value})")
        #amount raised
        if 'issue_size_number' in data_dict and 'face_value_num' in data_dict:
            amount_raised = calculate_amount_raised(data_dict)
            data_dict['amount_raised'] = amount_raised
            if debug:
                print(f"💰 Total Value: {amount_raised}")
        # Extract discount value number
        if 'Discount at which security is issued' in data_dict:
            discount_value_num, formatted_discount_value = extract_discount_value_number(data_dict['Discount at which security is issued'])
            data_dict['discount_value_num'] = discount_value_num
            data_dict['formatted_discount_value'] = formatted_discount_value
            if debug:
                print(f"💹 Extracted discount value: {discount_value_num}% ({formatted_discount_value})")
        # Extract issue price number
        if 'Issue Price' in data_dict:
            issue_price_num, formatted_issue_price = extract_issue_price_number(data_dict['Issue Price'])
            data_dict['issue_price_num'] = issue_price_num
            data_dict['formatted_issue_price'] = formatted_issue_price
            if debug:
                print(f"💲 Extracted issue price: {issue_price_num} ({formatted_issue_price})")
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
