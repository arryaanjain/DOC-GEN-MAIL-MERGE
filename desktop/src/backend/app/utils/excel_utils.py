import pandas as pd
from typing import Dict, List, Any

class ExcelWriter:
    @staticmethod
    def write_to_excel(data_dict: Dict, xlsx_file_path: str, debug_info: List[Dict] = None, debug: bool = True):
        """Handle Excel file writing and formatting"""
        with pd.ExcelWriter(xlsx_file_path, engine='xlsxwriter') as writer:
            # Create main data sheet
            df = pd.DataFrame([data_dict])
            df.to_excel(writer, sheet_name="Terms", index=False)
            
            # Format sheets
            workbook = writer.book
            ExcelWriter._format_main_sheet(df, writer, workbook)
            
            # Add debug sheet if requested
            if debug and debug_info:
                debug_df = ExcelWriter._add_debug_sheet(debug_info, writer, workbook)
                ExcelWriter._format_debug_sheet(debug_df, writer, workbook)

    @staticmethod
    def append_to_excel(data_to_append, xlsx_file_path: str, debug_info_to_append: list = None, debug: bool = True):
        # Read existing data
        try:
            with pd.ExcelFile(xlsx_file_path) as reader:
                if "Terms" in reader.sheet_names:
                    existing_df = pd.read_excel(reader, sheet_name="Terms")
                else:
                    existing_df = pd.DataFrame()
                if "Debug_Log" in reader.sheet_names:
                    existing_debug_df = pd.read_excel(reader, sheet_name="Debug_Log")
                else:
                    existing_debug_df = pd.DataFrame()
        except FileNotFoundError:
            existing_df = pd.DataFrame()
            existing_debug_df = pd.DataFrame()

        # Append main data (as a new row)
        if data_to_append:
            if existing_df.empty:
                # First write, use keys from dict or raise error for list
                if isinstance(data_to_append, dict):
                    new_df = pd.DataFrame([data_to_append])
                else:
                    raise ValueError("Cannot append list of values to empty sheet; need column names.")
            else:
                # Align dict to existing columns
                if isinstance(data_to_append, dict):
                    row = {col: data_to_append.get(col, "") for col in existing_df.columns}
                    new_df = pd.DataFrame([row])
                elif isinstance(data_to_append, list) and len(data_to_append) == len(existing_df.columns):
                    new_df = pd.DataFrame([data_to_append], columns=existing_df.columns)
                elif isinstance(data_to_append, list) and isinstance(data_to_append[0], dict):
                    rows = [{col: d.get(col, "") for col in existing_df.columns} for d in data_to_append]
                    new_df = pd.DataFrame(rows)
                elif isinstance(data_to_append, list) and isinstance(data_to_append[0], list):
                    new_df = pd.DataFrame(data_to_append, columns=existing_df.columns)
                else:
                    raise ValueError("data_to_append shape does not match existing columns")
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        else:
            combined_df = existing_df

        # Append debug info (as new rows)
        if debug and debug_info_to_append:
            if existing_debug_df.empty:
                new_debug_df = pd.DataFrame(debug_info_to_append)
            else:
                # Align each row to existing columns
                if isinstance(debug_info_to_append[0], dict):
                    rows = [{col: d.get(col, "") for col in existing_debug_df.columns} for d in debug_info_to_append]
                    new_debug_df = pd.DataFrame(rows)
                else:
                    new_debug_df = pd.DataFrame(debug_info_to_append, columns=existing_debug_df.columns)
            combined_debug_df = pd.concat([existing_debug_df, new_debug_df], ignore_index=True)
        else:
            combined_debug_df = existing_debug_df

        # Write back to Excel
        with pd.ExcelWriter(xlsx_file_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            combined_df.to_excel(writer, sheet_name="Terms", index=False)
            ExcelWriter._format_main_sheet(combined_df, writer, workbook)
            if debug and not combined_debug_df.empty:
                combined_debug_df.to_excel(writer, sheet_name="Debug_Log", index=False)
                ExcelWriter._format_debug_sheet(combined_debug_df, writer, workbook)

    @staticmethod
    def _add_debug_sheet(debug_info: List[Dict], writer, workbook) -> Any:
        """Add a debug sheet with detailed information"""
        if debug_info:
            debug_df = pd.DataFrame(debug_info)
            debug_df.to_excel(writer, sheet_name="Debug_Log", index=False)
            
            # Format the debug sheet
            worksheet = writer.sheets['Debug_Log']
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#F9CB9C',
                'border': 1
            })
            
            for col_num, column_name in enumerate(debug_df.columns.values):
                worksheet.write(0, col_num, column_name, header_format)
                max_length = max(
                    len(str(column_name)),
                    debug_df[column_name].astype(str).map(len).max() if not debug_df.empty else 0
                )
                worksheet.set_column(col_num, col_num, min(max_length + 2, 60))  # Cap at 60 characters
            return debug_df
        
    @staticmethod
    def _format_main_sheet(df, writer, workbook):
        """Format the main Terms sheet"""
        # ... existing Excel formatting code ...
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
        
    @staticmethod
    def _format_debug_sheet(debug_df, writer, workbook):
        """Format the debug sheet"""
        if not debug_df.empty:
            worksheet = writer.sheets['Debug_Log']
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#F9CB9C',
                'border': 1
            })
            
            for col_num, column_name in enumerate(debug_df.columns.values):
                worksheet.write(0, col_num, column_name, header_format)
                max_length = max(
                    len(str(column_name)),
                    debug_df[column_name].astype(str).map(len).max() if not debug_df.empty else 0
                )
                worksheet.set_column(col_num, col_num, min(max_length + 2, 60))
         