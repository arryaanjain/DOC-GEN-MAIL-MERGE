import os
from datetime import datetime

def get_output_directory(root_dir, processing_date):
    """Create and return the output directory path as root_dir/Month-YYYY/DD-MM-YYYY"""
    # Parse the date
    date_obj = datetime.strptime(processing_date, "%Y-%m-%d")
    month_year = date_obj.strftime("%B-%Y")      # e.g., June-2025
    day_month_year = date_obj.strftime("%d-%m-%Y")  # e.g., 28-06-2025

    # Build the full path
    month_dir = os.path.join(root_dir, month_year)
    day_dir = os.path.join(month_dir, day_month_year)

    # Create directories if they don't exist
    os.makedirs(day_dir, exist_ok=True)
    return day_dir