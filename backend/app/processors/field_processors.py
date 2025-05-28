from typing import Dict
from datetime import datetime
from app.utils.text_utils import extract_series_number, extract_issue_size_number
from app.utils.number_utils import extract_tenor_days_number, extract_face_value_number, extract_discount_value_number, extract_issue_price_number, calculate_amount_raised

class FieldProcessor:
    @staticmethod
    def process_special_fields(data_dict: Dict, debug: bool = True) -> Dict:
        """Process all special fields in the data dictionary"""
        processors = {
            'Product Code': FieldProcessor._process_product_code,
            'Issue Size': FieldProcessor._process_issue_size,
            'Tenor In Days': FieldProcessor._process_tenor_days,
            'Face Value': FieldProcessor._process_face_value,
            'Discount at which security is issued': FieldProcessor._process_discount_value,
            'Issue Price': FieldProcessor._process_issue_price,
            'Issue Opening Date': FieldProcessor._process_issue_opening_date,  # Add new processor
        }
        
        for field, processor in processors.items():
            if field in data_dict:
                processor(data_dict, debug)
        
         # Process date fields
        data_dict = FieldProcessor._process_date_fields(data_dict, debug)
        

        FieldProcessor._process_amount_raised(data_dict, debug)

        return data_dict

    @staticmethod
    def _process_product_code(data_dict: Dict, debug: bool):
        """Process Product Code field"""
        series_number = extract_series_number(data_dict['Product Code'])
        data_dict['series_number'] = series_number
        if debug:
            print(f"📎 Extracted series number: {series_number}")

    @staticmethod
    def _process_issue_size(data_dict: Dict, debug: bool):
        """Process Issue Size field"""
        issue_size, formatted_amount = extract_issue_size_number(data_dict['Issue Size'])
        data_dict['issue_size_num'] = issue_size
        data_dict['total_amount'] = formatted_amount

        if debug:
            print(f"📎 Extracted issue size: {issue_size}")
            print(f"💰 Extracted total amount: {formatted_amount}")
    @staticmethod
    def _process_tenor_days(data_dict: Dict, debug: bool):
        """Process Tenor In Days field"""
        tenor_days = extract_tenor_days_number(data_dict['Tenor In Days'])
        data_dict['tenor_days_num'] = tenor_days
        if debug:
            print(f"📎 Extracted tenor days: {tenor_days}")

    @staticmethod
    def _process_face_value(data_dict: Dict, debug: bool):
        """Process Face Value field"""
        face_value, formatted_face_value = extract_face_value_number(data_dict['Face Value'])
        data_dict['face_value_num'] = face_value
        data_dict['formatted_face_value'] = formatted_face_value
        if debug:
            print(f"💵 Extracted face value: {face_value} ({formatted_face_value})")

    @staticmethod
    def _process_amount_raised(data_dict: Dict, debug: bool):
        #amount raised
        if 'issue_size_num' in data_dict and 'face_value_num' in data_dict:
            amount_raised = calculate_amount_raised(data_dict)
            data_dict['amount_raised'] = amount_raised
            if debug:
                print(f"💰 Total Value: {amount_raised}")
        
    @staticmethod
    def _process_discount_value(data_dict: Dict, debug: bool):
        """Process Discount at which security is issued field"""
        discount_value_num, formatted_discount_value = extract_discount_value_number(data_dict['Discount at which security is issued'])
        data_dict['discount_value_num'] = discount_value_num
        data_dict['formatted_discount_value'] = formatted_discount_value
        if debug:
            print(f"💹 Extracted discount value: {discount_value_num}% ({formatted_discount_value})")

    @staticmethod
    def _process_issue_price(data_dict: Dict, debug: bool):
        """Process Issue Price field"""
        issue_price_num, formatted_issue_price = extract_issue_price_number(data_dict['Issue Price'])
        data_dict['issue_price_num'] = issue_price_num
        data_dict['formatted_issue_price'] = formatted_issue_price
        if debug:
            print(f"💲 Extracted issue price: {issue_price_num} ({formatted_issue_price})")

    @staticmethod
    def _process_date_fields(data_dict: Dict, debug: bool = True) -> Dict:
        """Process all date fields in the document"""
        date_fields_to_process = [
            "Issue Opening Date", "Issue Closing Date", "Pay-in-Date", 
            "Date of Allotment", "Deemed Date of Allotment", "Initial Fixing Date",
            "Final Fixing Date", "Redemption Date", "Coupon / Dividend payment dates"
        ]
        
        for date_field in date_fields_to_process:
            date_value = data_dict.get(date_field)
            if date_value:
                try:
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
                        
                        if debug:
                            print(f"📅 Processed {date_field}: {date_value} → {data_dict[date_key]}")
                
                except Exception as e:
                    if debug:
                        print(f"⚠️ Could not parse date field {date_field}: {date_value} - {e}")
                
                except Exception as e:
                    if debug:
                        print(f"⚠️ Could not parse date field {date_field}: {date_value} - {e}")
        
        return data_dict
    
    @staticmethod
    def _process_issue_opening_date(data_dict: Dict, debug: bool):
        """Process Issue Opening Date into individual digits"""
        date_value = data_dict.get('Issue Opening Date')
        if not date_value:
            return

        try:
            date_formats = ["%B %d, %Y", "%B %d,%Y", "%d %B %Y", "%d/%m/%Y", "%Y-%m-%d"]
            parsed_date = None
            
            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_value, fmt)
                    break
                except ValueError:
                    continue
            
            if parsed_date:
                formatted_date = parsed_date.strftime("%d%m%Y")
                # Create mapping for digit positions
                field_map = {0: 'd1', 1: 'd2', 2: 'm1', 3: 'm2', 
                           4: 'y1', 5: 'y2', 6: 'y3', 7: 'y4'}
                
                # Map each digit to its corresponding field
                for i, char in enumerate(formatted_date):
                    data_dict[field_map[i]] = char
                
                # Store formatted date
                data_dict['issue_opening_date_formatted'] = formatted_date
                
                if debug:
                    print(f"📅 Processed Issue Opening Date: {date_value} → {formatted_date}")
        
        except Exception as e:
            if debug:
                print(f"⚠️ Could not parse Issue Opening Date: {date_value} - {e}")

        