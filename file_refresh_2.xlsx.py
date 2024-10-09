import os
import logging
from win32com.client import Dispatch
import pythoncom
import time

# Configure logging
logging.basicConfig(filename='file_refresh_log.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

sharepoint_path = r'C:\Path\To\Directory'
files_to_skip = ['skipfile.xlsx']
directories_to_skip = ['SkipDirectory']

def refresh_excel_file(excel_app, file_path):
    try:
        workbook = excel_app.Workbooks.Open(Filename=file_path, UpdateLinks=3)
        excel_app.CalculateUntilAsyncQueriesDone()
        excel_app.DisplayAlerts = False  # Disable alerts to prevent most popup messages
        workbook.RefreshAll()
        workbook.Save()
        workbook.Close(SaveChanges=False)
        excel_app.DisplayAlerts = True  # Enable alerts after operations are done
        logging.info(f"Processed {file_path}")
    except Exception as e:
        logging.error(f"Failed to process {file_path}: {e}")
        try:
            workbook.Close(SaveChanges=False)  # Try to close without saving changes
        except:
            pass  # If workbook is already closed, pass
        excel_app.DisplayAlerts = True  # Ensure alerts are re-enabled for the next file

def main():
    pythoncom.CoInitialize()  # Initialize the COM library
    excel_app = Dispatch("Excel.Application")
    excel_app.Visible = False  # Run Excel in the background

    try:
        for root, dirs, files in os.walk(sharepoint_path):
            dirs[:] = [d for d in dirs if d not in directories_to_skip]
            for file in files:
                if file.endswith(".xlsx") and file not in files_to_skip:
                    file_path = os.path.join(root, file)
                    refresh_excel_file(excel_app, file_path)
    except Exception as e:
        logging.error(f"An error occurred: {e}")
    finally:
        excel_app.Quit()
        pythoncom.CoUninitialize()  # Uninitialize the COM library

if __name__ == "__main__":
    main()
