from src.backend.app.utils.text_utils import format_indian_number

def extract_tenor_days_number(tenor_text):
    """Extract number from Tenor in Days text"""
    import re
    
    if not tenor_text:
        return ""
    
    # Find numbers in the text
    numbers = re.findall(r'\d+', tenor_text)
    
    # Return first number found or empty string if none found
    return numbers[0] if numbers else ""

def extract_face_value_number(face_value):
    """
    Extract number from Face Value text and return both raw and formatted values
    Returns tuple: (number_as_string, formatted_with_commas)
    Example: "Rs. 1,00,000/- (Rupees One Lakh Only)" -> ("100000", "1,00,000")
    """
    import re
    
    if not face_value:
        return ("", "")
    
    # Remove commas from numbers
    cleaned_text = face_value.replace(',', '')
    
    # Find numbers including decimals
    numbers = re.findall(r'\d+\.?\d*', cleaned_text)
    
    if numbers:
        try:
            # Convert to integer if possible
            number_str = numbers[0]
            number_int = int(float(number_str))
            formatted = format_indian_number(number_int)
            return (str(number_int), formatted)
        except (ValueError, TypeError):
            return (numbers[0], numbers[0])
    
    return ("", "")

def calculate_amount_raised(data_dict):
    """
    Calculate total value by multiplying issue size and face value,
    converting from lakhs to crores
    """
    try:
        # Get values and convert to float
        issue_size_num = float(data_dict.get('issue_size_num', 0))
        face_value = float(data_dict.get('face_value_num', 0))
        print(f"Calculating amount raised: issue_size_num={issue_size_num}, face_value={face_value}")
        # Calculate total in lakhs
        total_in_lakhs = issue_size_num * face_value
        
        # Convert to crores (1 crore = 100 lakhs)
        total_in_crores = total_in_lakhs / 10000000
        
        # Format with comma separators and 2 decimal places
        formatted_value = f"Rs {total_in_crores:,.2f} crores"
        
        return formatted_value
    except (ValueError, TypeError) as e:
        print(f"⚠️ Error calculating total value: {e}")
        return "Rs 0 crores"

def extract_discount_value_number(discount_text):
    """
    Extract first complete number from discount text, handling commas.
    Returns a tuple: (number_as_string, formatted_with_commas)
    Example: "Rs. 4,500/- (Rupees Four Thousand Five Hundred Only)" -> ("4500", "4,500")
    """
    import re

    if not discount_text:
        return ("", "")

    # First find any number pattern that may include commas
    number_pattern = r'(?:Rs\.?\s*)?(\d+(?:,\d+)*(?:\.\d+)?)'
    matches = re.search(number_pattern, discount_text)

    if matches:
        # Remove commas from the matched number
        number_str = matches.group(1).replace(',', '')
        try:
            # Try to convert to int if possible, else float
            if '.' in number_str:
                number_val = float(number_str)
            else:
                number_val = int(number_str)
            formatted = format_indian_number(int(float(number_str)))
            return (number_str, formatted)
        except (ValueError, TypeError):
            return (number_str, number_str)

    return ("", "")

def extract_issue_price_number(issue_price_text):
    """
    Extract first complete number from Issue Price text, handling commas.
    Returns a tuple: (number_as_string, formatted_with_commas)
    Example: "Rs. 95,500/- (Rupees Ninety-Five Thousand Five Hundred Only) per Debenture"
    Output: ("95500", "95,500")
    """
    import re

    if not issue_price_text:
        return ("", "")

    # Find first number pattern that may include commas
    number_pattern = r'(?:Rs\.?\s*)?(\d+(?:,\d+)*(?:\.\d+)?)'
    match = re.search(number_pattern, issue_price_text)
    if match:
        number_str = match.group(1).replace(',', '')
        try:
            number_int = int(float(number_str))
            formatted = format_indian_number(number_int)
            return (number_str, formatted)
        except (ValueError, TypeError):
            return (number_str, number_str)
    return ("", "")
