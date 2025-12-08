import sys
import os
import tkinter as tk

# Add the directory to path to import excel_merger
sys.path.append(os.path.join(os.getcwd(), 'analytics_ui'))

try:
    from analytics_ui.excel_merger import ExcelMerger
    print("Successfully imported ExcelMerger")
    
    # Check if method exists
    if hasattr(ExcelMerger, 'open_range_editor'):
        print("Method 'open_range_editor' exists in ExcelMerger class.")
    else:
        print("ERROR: Method 'open_range_editor' NOT found in ExcelMerger class.")
        sys.exit(1)
        
    print("Verification script finished successfully.")
    
except ImportError as e:
    print(f"ImportError: {e}")
    sys.exit(1)
except Exception as e:
    print(f"An error occurred: {e}")
    sys.exit(1)
