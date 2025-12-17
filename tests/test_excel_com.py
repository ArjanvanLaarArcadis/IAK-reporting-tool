# IAK Reporting Tool - Excel COM test
# Copyright (C) 2024-2025 Arcadis Nederland B.V.
#
# SPDX-License-Identifier: GPL-3.0-or-later
# See LICENSE file for full license text.

"""
Test script to check if Excel COM interface is working properly.
Run this script to diagnose Excel COM connectivity issues.
"""

import sys

try:
    import win32com.client
    import pythoncom
    print("✓ win32com.client imported successfully")
except ImportError as e:
    print(f"✗ Failed to import win32com.client: {e}")
    print("Solution: Install pywin32 package with: pip install pywin32")
    sys.exit(1)

def test_excel_com():
    """Test Excel COM interface connectivity."""
    excel_app = None
    try:
        print("Testing Excel COM interface...")
        
        # Initialize COM
        pythoncom.CoInitialize()
        print("✓ COM initialized successfully")
        
        # Try to create Excel application
        excel_app = win32com.client.Dispatch("Excel.Application")
        print("✓ Excel.Application created successfully")
        
        # Test basic Excel functionality
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        print("✓ Excel configured successfully")
        
        # Try to create a new workbook
        wb = excel_app.Workbooks.Add()
        print("✓ New workbook created successfully")
        
        # Close the workbook
        wb.Close(SaveChanges=False)
        print("✓ Workbook closed successfully")
        
        print("\n🎉 Excel COM interface is working properly!")
        return True
        
    except Exception as e:
        print(f"\n❌ Excel COM test failed: {e}")
        print(f"Error type: {type(e).__name__}")
        
        if "ConnectionRefusedError" in str(type(e)) or "10061" in str(e):
            print("\n🔧 Troubleshooting steps for ConnectionRefusedError:")
            print("1. Make sure Microsoft Excel is installed")
            print("2. Run: regsvr32 /i:user excel.exe")
            print("3. Run as Administrator: regsvr32 excel.exe")
            print("4. Restart Windows")
            print("5. Check if Excel can open manually")
            
        elif "Permission" in str(e) or "Access" in str(e):
            print("\n🔧 Troubleshooting steps for Permission errors:")
            print("1. Run Python script as Administrator")
            print("2. Check DCOM configuration for Excel")
            print("3. Enable 'Interactive User' in DCOM settings")
            
        return False
        
    finally:
        # Clean up
        if excel_app:
            try:
                excel_app.Quit()
                print("✓ Excel application closed")
            except:
                pass
        
        try:
            pythoncom.CoUninitialize()
            print("✓ COM uninitialized")
        except:
            pass

if __name__ == "__main__":
    print("Excel COM Interface Test")
    print("=" * 30)
    success = test_excel_com()
    
    if not success:
        print("\n💡 If Excel COM is not working, the PDF export feature will fail.")
        print("💡 You can still generate Excel reports, just skip the PDF export step.")
