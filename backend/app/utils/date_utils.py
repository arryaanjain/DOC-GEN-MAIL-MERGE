from datetime import datetime

def handle_processing_date(date_string):
    """
    Process the date received from frontend and format it for Excel output
    Args:
        date_string (str): Date in YYYY-MM-DD format from frontend
    Returns:
        dict: Dictionary containing formatted date components
    """
    
    try:
        # Parse the date string (expected format: YYYY-MM-DD)
        parsed_date = datetime.strptime(date_string, "%Y-%m-%d")
        
        # Format into DDMMYYYY
        formatted_date = parsed_date.strftime("%d%m%Y")

        formatted_string = parsed_date.strftime("%d %B, %Y")
        
        # Create dictionary with date components
        date_components = {
            'processing_date': date_string,
            'processing_date_formatted': formatted_date,
            'processing_date_string': formatted_string,
            'D1': formatted_date[0],
            'D2': formatted_date[1],
            'M1': formatted_date[2],
            'M2': formatted_date[3],
            'Y1': formatted_date[4],
            'Y2': formatted_date[5],
            'Y3': formatted_date[6],
            'Y4': formatted_date[7]
        }
        
        return date_components
    except ValueError as e:
        print(f"⚠️ Error processing date: {e}")
        return None
    