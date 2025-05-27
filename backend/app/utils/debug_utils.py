from typing import List, Dict
import re
class DebugLogger:
    @staticmethod
    def print_row_debug(row_idx: int, cells: list):
        """Print debug information for a row"""
        # Extract text from cells if they are _Cell objects
        cell_texts = []
        for cell in cells:
            if hasattr(cell, 'text'):  # Check if it's a _Cell object
                cell_text = cell.text.strip().replace('\n', ' ').replace('\t', ' ')
                cell_text = re.sub(r'\s+', ' ', cell_text).strip()
                cell_texts.append(cell_text)
            else:
                cell_texts.append(str(cell).strip())

        print(f"  Row {row_idx + 1}: {len(cell_texts)} columns, {len([c for c in cell_texts if c])} non-empty")
        for i, cell_text in enumerate(cell_texts):
            if cell_text:
                print(f"    Col {i+1}: '{cell_text[:50]}{'...' if len(cell_text) > 50 else ''}'")
    
    @staticmethod
    def add_debug_info(table_idx: int, row_idx: int, method: str, key: str, value: str, cells: list) -> dict:
        """Create debug information entry"""
        return {
            'Table': table_idx + 1,
            'Row': row_idx + 1,
            'Method': method,
            'Key': key,
            'Value': value[:100] + '...' if len(value) > 100 else value,
            'Original_Cells': ' | '.join([f"Col{i+1}: {cell}" for i, cell in enumerate(cells) if cell.strip()])
        }

    @staticmethod
    def log_processing_date(processing_date: str):
        """Log processing date addition"""
        print(f"📅 Added processing date: {processing_date}")

    @staticmethod
    def log_strategy_success(strategy_name: str, key: str, value: str):
        """Log successful strategy extraction"""
        print(f"    ✓ {strategy_name}: Added {key} = {value[:50]}...")