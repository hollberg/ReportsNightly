"""
refresh_xlsx.py

Given a directory that contains Excel files
    * Loop over all files and subdirectories to find all *.xlsx files
    * For each file:
        * Open
        * Data -> Refresh All
        * Save and Close
"""


import os
import logging
from win32com.client import Dispatch
import time
import pythoncom

# Configure logging
logging.basicConfig(filename='file_refresh_log.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

sharepoint_path = fr'C:\Users\mitchell.hollberg\Community Foundation for Greater Atlanta\IT - Documents\Public\Reports'

files_to_skip = [
    'Finance - GL Balance and GL Activity.xlsx',
    'Grants Report.xlsx',
    'zAdmin - CSuite Data Quality.xlsx',
]

directories_to_skip = [
    'Data Quality',
]


def refresh_excel_file(excel_app, file_path):

    try:
        start_time = time.time()
        logging.info(f"Opening {file_path}...")

        workbook = excel_app.Workbooks.Open(os.path.abspath(file_path), UpdateLinks=0)

        logging.info(f"Refreshing data in {os.path.basename(file_path)}...")
        workbook.RefreshAll()
        excel_app.CalculateUntilAsyncQueriesDone()

        workbook.Save()
        workbook.Close()

        end_time = time.time()
        duration = end_time - start_time
        logging.info(f"Successfully refreshed {os.path.basename(file_path)} in {duration} seconds.")
        print(f"Successfully refreshed {os.path.basename(file_path)} in {duration} seconds.")

    except Exception as e:
        logging.error(f"Failed to refresh {os.path.basename(file_path)}: {str(e)}")
        print(f"Failed to refresh {os.path.basename(file_path)}: {str(e)}")
        if 'workbook' in locals():  # Check if 'workbook' is defined
            try:
                workbook.Close(False)  # Attempt to close the workbook without saving
            except Exception as close_error:
                logging.error(f"Failed to close {os.path.basename(file_path)}: {str(close_error)}")
                print(f"Failed to close {os.path.basename(file_path)}: {str(close_error)}")



def find_xlsx_files(starting_directory, files_to_skip, skip_directories):
    for root, dirs, files in os.walk(starting_directory, topdown=True):
        dirs[:] = [d for d in dirs if d not in skip_directories]
        for file in files:
            if file.endswith(".xlsx") and file not in files_to_skip:
                file_path = os.path.join(root, file)
                yield file_path


def main():
    pythoncom.CoInitialize()  # Initialize COM library at the beginning
    excel_app = Dispatch("Excel.Application")
    excel_app.Visible = True
    excel_app.DisplayAlerts = True

    for file_path in find_xlsx_files(sharepoint_path, files_to_skip, directories_to_skip):
        refresh_excel_file(excel_app, file_path)

    excel_app.Quit()
    pythoncom.CoUninitialize()  # Uninitialize COM library at the end
    logging.info("All files processed successfully.")


if __name__ == "__main__":
    main()
