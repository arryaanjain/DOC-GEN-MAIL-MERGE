from typing import Dict, List, Tuple

class ExtractionStrategy:
    @staticmethod
    def apply_strategy1(cells: List[str], valid_headers: List[str], skip_phrases: List[str]) -> Tuple[str, str, str]:
        """Strategy 1: Find key-value pairs in adjacent columns"""
        num_cols = len(cells)
        key = value = extraction_method = ""
        
        if num_cols >= 2:
            for i in range(num_cols - 1):
                if cells[i].strip():
                    cell_lower = cells[i].lower()
                    if (not any(phrase in cell_lower for phrase in skip_phrases) or 
                        any(header in cell_lower for header in valid_headers)):
                        key = cells[i].strip()
                        if num_cols > 2 and cells[1].strip() and cells[2].strip():
                            value = f"{cells[1].strip()} | {cells[2].strip()}" if cells[1].strip() != cells[2].strip() else cells[1].strip()
                        else:
                            value = cells[i+1].strip()
                        extraction_method = f"Strategy1_Col{i+1}→Col{i+2}"
                        break
        
        return key, value, extraction_method
    
    @staticmethod
    def apply_strategy2(cells: List[str], valid_headers: List[str], skip_phrases: List[str]) -> Tuple[str, str, str]:
        """Strategy 2: Handle single column with meaningful content"""
        num_cols = len(cells)
        key = value = extraction_method = ""
        if not key and num_cols >= 1:
            for i, cell in enumerate(cells):
                cell_lower = cell.strip().lower()
                if (cell.strip() and 
                    not any(skip in cell_lower for skip in skip_phrases) or
                    any(header in cell_lower for header in valid_headers)):
                    if len(cell.strip()) > 3:  # Avoid very short strings
                        key = cell.strip()  # Use actual content instead of Field_X_Y_Z
                        value = cell.strip()
                        extraction_method = f"Single_Col{i+1}"
        return key, value, extraction_method
    
    @staticmethod
    def apply_strategy3(cells: List[str], data_dict) -> Tuple[str, str, str]:
        """Strategy 3: Handle multi-column data (3+ columns) which is different"""
        num_cols = len(cells)
        key = value = extraction_method = ""
        if not key and num_cols >= 3:
            # Check if this is a Coupon row
            first_cell = cells[0].strip().lower() if cells[0] else ""
            
            if first_cell == 'coupon':
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
                        
                # Set method for debug info
                extraction_method = "Coupon_Special"
                
        return key, value, extraction_method