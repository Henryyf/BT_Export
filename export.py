import os
import time
import datetime
import psutil
import pythoncom
import win32com.client


def save_excel_snapshot_fixed_range(
    pid=25308,
    iterations=10,
    interval_sec=10,
    out_dir=os.path.join(os.path.expanduser("~"), "Desktop", "export", "data"),
    all_sheets=False,        # set True to export A1:J100 from every sheet
    sheet_name=None          # set to a specific sheet name to force that sheet
):
    # 1) Verify PID
    try:
        proc = psutil.Process(pid)
    except psutil.NoSuchProcess:
        raise RuntimeError(f"No process with PID {pid} was found.")
    if "EXCEL" not in proc.name().upper():
        raise RuntimeError(f"PID {pid} is not an Excel process.")

    # 2) Attach to running Excel instance
    pythoncom.CoInitialize()
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        raise RuntimeError("Could not attach to running Excel instance.") from e

    if xl.Workbooks.Count == 0:
        raise RuntimeError("No workbook is currently open in Excel.")
    wb = xl.ActiveWorkbook

    # 3) Prepare output
    os.makedirs(out_dir, exist_ok=True)
    xl.DisplayAlerts = False

    # Helper to export a single sheet's A1:J100 as CSV (values only)
    def export_sheet_range(sh):
        # Fixed range A1:J100 regardless of UsedRange/headers/formulas
        src = sh.Range("A1:J100")
        # Create a temporary workbook and copy immediate values only
        temp_wb = xl.Workbooks.Add()
        dest = temp_wb.ActiveSheet.Range("A1").Resize(src.Rows.Count, src.Columns.Count)
        dest.Value2 = src.Value2  # immediate values (numeric/text), no formulas, no formats

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base = (sheet_name or sh.Name).replace(os.sep, "_")
        csv_path = os.path.join(out_dir, f"{base}_{ts}.csv")

        # Save as CSV (try UTF-8 first, then fallback to standard CSV)
        # xlCSVUTF8 = 62, xlCSV = 6
        try:
            temp_wb.SaveAs(csv_path, FileFormat=62)  # UTF-8 if supported
        except Exception:
            temp_wb.SaveAs(csv_path, FileFormat=6)   # ANSI fallback
        finally:
            temp_wb.Close(SaveChanges=False)

        print(f"Saved {csv_path}")

    # Decide which sheets to export
    def sheets_to_export():
        if sheet_name:
            try:
                return [wb.Sheets(sheet_name)]
            except Exception:
                raise RuntimeError(f"Sheet '{sheet_name}' not found in workbook '{wb.Name}'.")
        return list(wb.Sheets) if all_sheets else [wb.ActiveSheet]

    for i in range(iterations):
        for sh in sheets_to_export():
            export_sheet_range(sh)

        if i < iterations - 1:
            time.sleep(interval_sec)


if __name__ == "__main__":
    # Example: active sheet only, 10 iterations, every 10 seconds
    save_excel_snapshot_fixed_range(
        pid=25308,
        iterations=10,
        interval_sec=10,
        out_dir=os.path.join(os.path.expanduser("~"), "Desktop", "export", "data"),
        all_sheets=False,      # set True if you want every sheet exported each time
        sheet_name=None        # or set like "Sheet1" to force a particular sheet
    )
