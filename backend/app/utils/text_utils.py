# Add this function near the top of the file
def extract_series_number(product_code):
    """Extract series number from product code"""
    if not product_code:
        return ""
    
    # Split by spaces and get the last part
    parts = product_code.strip().split()
    if parts:
        last_part = parts[-1]
        # If it contains 'series' (case insensitive), return the whole last part
        if 'series' in last_part.lower():
            return last_part
        # If the last part is just a number, add 'Series' prefix
        if last_part.isdigit():
            return f"Series {last_part}"
    return ""

def extract_issue_size_number(issue_size):
    """
    Extract both the first number (issue size) and fourth number (total amount) from issue size text
    Returns tuple: (issue_size_num, formatted_total_amount)
    Example input: "75 Debentures bearing face value of Rs. 1,00,000/- each, issued at Rs.95,500/-, 
                total amounting to Rs.71,62,500/-"
    Example output: ("75", "71,62,500")
    """
    import re
    
    if not issue_size:
        return ("", "")
    
    # Remove all commas first
    cleaned_text = issue_size.replace(',', '')
    
    # Find all numbers (including decimals) in the text
    numbers = re.findall(r'\d+\.?\d*', cleaned_text)
    
    # Get first number (issue size)
    issue_size_num = numbers[0] if numbers else ""
    
    # Get fourth number if available and format it
    formatted_amount = ""
    if len(numbers) >= 4:
        total_amount = numbers[3]
        try:
            amount = int(float(total_amount))
            formatted_amount = format_indian_number(amount)
        except (ValueError, TypeError):
            pass
    
    return (issue_size_num, formatted_amount)

def format_indian_number(number):
    """Helper function to format number with Indian style commas"""
    str_num = str(number)
    length = len(str_num)
    
    if length <= 3:
        return str_num
    
    # Split the number into parts
    last_three = str_num[-3:]
    remaining = str_num[:-3]
    
    # Add commas every 2 digits in the remaining part
    formatted = ''
    for i, digit in enumerate(reversed(remaining)):
        if i % 2 == 1 and i != len(remaining) - 1:
            formatted = ',' + digit + formatted
        else:
            formatted = digit + formatted
    
    return formatted + ',' + last_three

