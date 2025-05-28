from docx import Document
import re
from app.utils.date_utils import handle_processing_date
from app.utils.excel_utils import ExcelWriter
from app.utils.debug_utils import DebugLogger
from app.core.extraction_strategies import ExtractionStrategy
from app.processors.field_processors import FieldProcessor

class DocumentProcessor:
    def __init__(self, debug=True):
        self.debug = debug
        self.data_dict = {}
        self.debug_info = []
        if self.debug:
            self.logger = DebugLogger()

    # Enhanced function for better table structure handling
    def convert_docx_to_xlsx(self, docx_file_path, xlsx_file_path, processing_date=None):

        doc = Document(docx_file_path)
        # Process tables
        self._process_tables(doc, processing_date)
        
        if processing_date:
            date_components = handle_processing_date(processing_date)
            if date_components:
                self.data_dict.update(date_components)
                if self.debug:
                    self.logger.log_processing_date(processing_date)
                                
                        
        # Process all special fields including dates
        self.data_dict = FieldProcessor.process_special_fields(self.data_dict, self.debug)
        ExcelWriter.write_to_excel(self.data_dict, xlsx_file_path, self.debug_info, self.debug)
        return self.data_dict
    
    def _process_tables(self, doc, processing_date):
        """Process all tables in the document"""
        # ... table processing logic using ExtractionStrategy ...
        for table_idx, table in enumerate(doc.tables):
            if self.debug:
                print(f"\n📋 Processing Table {table_idx + 1}:")
            
            for row_idx, row in enumerate(table.rows):
                # Clean and extract all cell contents
                cells = []
                for cell in row.cells:
                    cell_text = cell.text.strip().replace('\n', ' ').replace('\t', ' ')
                    cell_text = re.sub(r'\s+', ' ', cell_text).strip()
                    cells.append(cell_text)
                
                # Print debug info after cell text extraction
                if self.debug:
                    self.logger.print_row_debug(row_idx, cells)

                # Skip completely empty rows
                if not any(cells):
                    continue
                
                # Skip header rows (containing "TERMS OF ISSUE")
                if any("TERMS OF ISSUE" in cell.upper() for cell in cells):
                    continue
                
                # Define common phrases and headers
                skip_phrases = [
                    'terms of issue',
                    'terms and conditions',
                    '>>',
                    '***',
                    '*****'
                ]
                valid_headers = [
                    'issue price',
                    'issue opening date',
                    'issue closing date',
                    'discount at which security is issued',
                    'coupon'
                ]
                
                # Try strategies in order - stop when one succeeds
                strategies = [
                    (ExtractionStrategy.apply_strategy1, "Strategy 1"),
                    (ExtractionStrategy.apply_strategy2, "Strategy 2"),
                    (ExtractionStrategy.apply_strategy3, "Strategy 3")
                ]
                
                for strategy_func, strategy_name in strategies:
                    key, value, extraction_method = strategy_func(cells, valid_headers, skip_phrases) if strategy_name != "Strategy 3" else strategy_func(cells, self.data_dict)
                    
                    if key and key.strip():
                        clean_key = key.strip()
                        clean_value = value.strip() if value else "Present"
                        self.logger.log_strategy_success(strategy_name, clean_key, clean_value) if self.debug else None
                        # Handle duplicate keys
                        clean_key = self._handle_duplicate_key(clean_key)
                        
                        # Store the data
                        self.data_dict[clean_key] = clean_value
                        
                        # Add debug info
                        if self.debug:
                            self.debug_info.append(self.logger.add_debug_info(
                                table_idx, row_idx, strategy_name, clean_key, clean_value, cells
                            ))
                        # Break after first successful strategy
                        break

        if processing_date:
            self._handle_processing_date(processing_date)

    def _handle_duplicate_key(self, key: str) -> str:
        """Handle duplicate keys by adding counter"""
        original_key = key
        counter = 1
        while key in self.data_dict:
            key = f"{original_key}_{counter}"
            counter += 1
        return key

    def _handle_processing_date(self, processing_date: str):
        """Handle processing date components"""
        date_components = handle_processing_date(processing_date)
        if date_components:
            self.data_dict.update(date_components)
            if self.debug:
                print(f"📅 Added processing date: {processing_date}")