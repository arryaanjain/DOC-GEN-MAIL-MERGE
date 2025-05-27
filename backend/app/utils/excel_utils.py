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
         