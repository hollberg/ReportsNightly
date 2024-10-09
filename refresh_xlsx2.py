import os
import logging
from win32com.client import Dispatch
import win32gui
import win32con
import pythoncom
import time

# Configure logging
logging.basicConfig(filename='file_refresh_log.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

sharepoint_path = fr'C:\Users\mitchell.hollberg\Community Foundation for Greater Atlanta\IT - Documents\Public\Reports'


files_to_skip = [
    'Finance - GL Balance and GL Activity.xlsx',
    'PO Portfolio Report - CS and RE Data - BETA.xlsx',
    # 'Actions Report - Raisers Edge.xlsx',
    # 'zAdmin - CSuite Data Quality.xlsx',
    'zAdmin - Current Staff Active Directory Export.xlsx',
    'zAdmin - Custom Reports.xlsx'
]

directories_to_skip = [
    'Data Quality',
]

def refresh_excel_file(excel_app, file_path):
    try:
        workbook = excel_app.Workbooks.Open(Filename=file_path, UpdateLinks=3)
        time.sleep(10)
        workbook.Save()
        excel_app.CalculateUntilAsyncQueriesDone()
        excel_app.DisplayAlerts = True  # Disable alerts to prevent most popup messages
        workbook.RefreshAll()

        # Wait for the refresh to complete (check for connections if you use them)
        for conn in workbook.Connections:
            # Disable background queries to address "connecting to server" issue
            # conn.OLEDBConnection.BackgroundQuery = False
            if conn.OLEDBConnection and conn.OLEDBConnection.BackgroundQuery:
                while conn.OLEDBConnection.Refreshing:
                    time.sleep(1)

        workbook.Save()
        workbook.Close(SaveChanges=False)
        excel_app.DisplayAlerts = True  # Enable alerts after operations are done
        logging.info(f"Processed {file_path}")
        print(f"Processed {file_path}")
    except Exception as e:
        logging.error(f"Failed to process {file_path}: {e}")
        try:    # May be due to excel "Trying to contact server" message; pass <ESC>
            resolve_excel_contact_server_message()

            workbook.Close(SaveChanges=False)  # Try to close without saving changes
        except Exception as e:
            pass  # If workbook is already closed, pass
    excel_app.DisplayAlerts = True  # Ensure alerts are re-enabled for the next file


def resolve_excel_contact_server_message(excel_app):
    """
    Sometimes Excel workbook opens and gives a message:
    "Contacting the server for more information. Press ESC to cancel."
    This locks workbook interactivity and breaks the script. When this happens,
    pass "ESC" to the Excel instance and see if that fixes the blockage.
    :return:
    """
    print(excel_app.StatusBar)
    excel_app.SendKeys("{ESC}")
    # Get the handle of the Excel window
    # excel_handle = win32gui.FindWindow("XLMAIN", None)
    # # Send <ESC>
    # win32gui.PostMessage(excel_handle, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
    # time.sleep(0.1)
    # win32gui.PostMessage(excel_handle, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
    time.sleep(1)


def close_excel_with_retry(excel_app, retries=5, delay=1):
    for attempt in range(retries):
        try:
            excel_app.Quit()
            return
        except pythoncom.com_error as e:
            logging.error(f"Failed to quit Excel on attempt {attempt + 1}: {e}")
            time.sleep(delay)
    logging.error("Failed to quit Excel after multiple attempts")

def main():
    pythoncom.CoInitialize()  # Initialize the COM library
    excel_app = Dispatch("Excel.Application")
    excel_app.Visible = True  # Run Excel in the foreground

    for root, dirs, files in os.walk(sharepoint_path):
        dirs[:] = [d for d in dirs if d not in directories_to_skip]
        for file in files:
            if file.endswith(".xlsx") and file not in files_to_skip:
                file_path = os.path.join(root, file)
                try:
                    refresh_excel_file(excel_app, file_path)
                except Exception as e:
                    logging.error(f"An error occurred: {e}")
                    print(f'Error with file {root}/{file}')
                    resolve_excel_contact_server_message(excel_app)
                    refresh_excel_file(excel_app, file_path)

    close_excel_with_retry(excel_app)
    pythoncom.CoUninitialize()  # Uninitialize the COM library

if __name__ == "__main__":
    main()
