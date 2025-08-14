import os
import time
import datetime
import psutil
import pythoncom
import win32com.client


def save_excel_snapshot(pid=25308, iterations=10, interval_sec=10):
    # Verify PID exists
    try:
        proc = psutil.Process(pid)
    except psutil.NoSuchProcess:
        raise RuntimeError(f"No process with PID {pid} was found.")

    # Get the main window title (should contain workbook name)
    title = proc.name()  # just the process name
    if "EXCEL" not in title.upper():
        raise RuntimeError(f"PID {pid} is not an Excel process.")

    # Attach to Excel COM instance
    pythoncom.CoInitialize()
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        raise RuntimeError("Could not attach to running Excel instance.") from e

    # Try to guess the workbook
    # Note: The PID-to-instance mapping is tricky; this assumes one Excel instance
    if xl.Workbooks.Count == 0:
        raise RuntimeError("No workbook is currently open in Excel.")
    wb = xl.ActiveWorkbook
    sheet = wb.ActiveSheet

    desktop = os.path.join(os.path.expanduser("~"), "Desktop", "export")

    for i in range(iterations):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = os.path.join(desktop, f"excel{timestamp}.csv")

        # Copy values to a temporary workbook to avoid touching original
        src = sheet.UsedRange
        temp_wb = xl.Workbooks.Add()
        dest = temp_wb.ActiveSheet.Range("A1").Resize(src.Rows.Count, src.Columns.Count)
        dest.Value = src.Value  # values only

        # xlCSV = 6
        temp_wb.SaveAs(csv_path, FileFormat=6)
        temp_wb.Close(SaveChanges=False)

        print(f"Saved {csv_path}")
        if i < iterations - 1:
            time.sleep(interval_sec)


if __name__ == "__main__":
    save_excel_snapshot(pid=25308, iterations=10, interval_sec=10)
