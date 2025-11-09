import os
import pythoncom
import win32com.client
from pathlib import Path
import time

def password_protect_excel_com(folder_path, password):
    """
    Use Excel COM to add real open password protection
    """
    folder_path = Path(folder_path)
    excel_files = list(folder_path.rglob("*.xlsx")) + list(folder_path.rglob("*.xls"))
    
    # Initialize COM
    pythoncom.CoInitialize()
    
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    
    processed = 0
    failed = 0
    
    for file_path in excel_files:
        print(f"Processing: {file_path.name}")
        
        try:
            # Open workbook
            workbook = excel_app.Workbooks.Open(str(file_path))
            
            # Set password to open
            workbook.Password = password
            
            # Save and close
            workbook.Save()
            workbook.Close()
            
            processed += 1
            print(f"✓ Successfully password protected: {file_path.name}")
            
        except Exception as e:
            failed += 1
            print(f"✗ Failed: {file_path.name}")
            print(f"  Error: {str(e)}")
        
        time.sleep(1)  # Give Excel time to process
    
    excel_app.Quit()
    pythoncom.CoUninitialize()
    
    print(f"\nSUMMARY: {processed} files protected, {failed} files failed")

# Usage
folder_path = input("Enter folder path: ").strip()
password = input("Enter open password: ").strip()
password_protect_excel_com(folder_path, password)