import os
import sys
import shutil
import traceback
from datetime import datetime

import win32com.client as win32

MAIN_REPORT_NAME = ""      
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb")
MAKE_BACKUP = True                        

JOBS = [
    # ---------------- 1. Top 15 ----------------
    {
        "name": "Top 15",
        "website_file": "1 Top15",
        "website_sheet": None,
        "remove_blank_rows": True,
        "shift": {
            "start_row": 6,
            "from_cols": ("D", "I"),
            "left_by": 1,
        },
        "copy": {
            "start_row": 9,
            "cols": ("C", "H"),
            "skip_merged_rows": False,
        },
        "target_sheet": "d_Top 15",
        "target_cell": "A5",
        "clear_old_data": True,
    },

    # ---------------- 2. ----------------
    # Copy the block below, remove the #, and fill in the values.
    # {
    #     "name": "",
    #     "website_file": "2 ",
    #     "website_sheet": None,
    #     "remove_blank_rows": True,
    #     "shift": None,
    #     "copy": {
    #         "start_row": 0,
    #         "cols": ("A", "A"),
    #         "skip_merged_rows": False,
    #     },
    #     "target_sheet": "",
    #     "target_cell": "A5",
    #     "clear_old_data": True,
    # },

    # ---------------- 3. ----------------

    # ---------------- 4. ----------------

    # ---------------- 5. ----------------

    # ---------------- 6. ----------------

    # ---------------- 7. ----------------

    # ---------------- 8. ----------------

    # ---------------- 9. ----------------
]


# =====================================================================
# HELPERS
# =====================================================================

LOG_LINES = []


def log(msg=""):
    print(msg)
    LOG_LINES.append(msg)


def base_folder():
    """Folder of the .exe (or the .py when run directly)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def col_num(letter):
    """'A' -> 1, 'H' -> 8, 'AA' -> 27"""
    n = 0
    for ch in letter.upper().strip():
        n = n * 26 + (ord(ch) - 64)
    return n


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def find_file(folder, name):
    """Find an Excel file whose name starts with `name`. Exact match wins, else newest."""
    matches = []
    for f in os.listdir(folder):
        if f.startswith("~$"):          # Excel temp/lock files
            continue
        stem, ext = os.path.splitext(f)
        if ext.lower() in EXCEL_EXTS and stem.lower().startswith(name.lower()):
            matches.append(os.path.join(folder, f))

    if not matches:
        raise FileNotFoundError(f"No Excel file starting with '{name}' found in {folder}")

    exact = [m for m in matches if os.path.splitext(os.path.basename(m))[0].lower() == name.lower()]
    if exact:
        return exact[0]
    return max(matches, key=os.path.getmtime)


def backup_report(path):
    folder = os.path.join(os.path.dirname(path), "Backup")
    os.makedirs(folder, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(folder, f"{stem}_{stamp}{ext}")
    shutil.copy2(path, dest)
    return dest


# =====================================================================
# READ + CLEAN website SHEET (all done in memory, website file is not changed)
# =====================================================================

def read_sheet(ws):
    """Return list of rows: {'values': [...], 'merged': bool}, plus last column number."""
    ur = ws.UsedRange
    last_row = ur.Row + ur.Rows.Count - 1
    last_col = ur.Column + ur.Columns.Count - 1

    data = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col)).Value
    if not isinstance(data, tuple):          # single cell sheet
        data = ((data,),)

    rows = []
    for r in range(1, last_row + 1):
        # MergeCells returns True / False / None (None = some cells merged)
        merged = ws.Range(ws.Cells(r, 1), ws.Cells(r, last_col)).MergeCells
        rows.append({"values": list(data[r - 1]), "merged": merged is not False})
    return rows, last_col


def transform(rows, last_col, job):
    # Make every row wide enough for the columns we use
    needed = last_col
    if job.get("shift"):
        needed = max(needed, col_num(job["shift"]["from_cols"][1]))
    needed = max(needed, col_num(job["copy"]["cols"][1]))
    for row in rows:
        row["values"] += [None] * (needed - len(row["values"]))

    # 1. Remove fully empty rows
    if job.get("remove_blank_rows"):
        rows = [r for r in rows if not all(is_blank(v) for v in r["values"])]

    # 2. Shift data left (merged rows are block boundaries and are left as they are)
    sh = job.get("shift")
    if sh:
        c1 = col_num(sh["from_cols"][0]) - 1
        c2 = col_num(sh["from_cols"][1]) - 1
        n = sh["left_by"]
        for i in range(sh["start_row"] - 1, len(rows)):
            row = rows[i]
            if row["merged"]:
                continue
            v = row["values"]
            moved = v[c1:c2 + 1]
            for k, val in enumerate(moved):
                v[c1 - n + k] = val
            for k in range(c2 - n + 1, c2 + 1):
                v[k] = None

    # 3. Pick the rows/columns to copy
    cp = job["copy"]
    a = col_num(cp["cols"][0]) - 1
    b = col_num(cp["cols"][1]) - 1
    out = []
    for i in range(cp["start_row"] - 1, len(rows)):
        row = rows[i]
        if all(is_blank(v) for v in row["values"]):
            break
        if cp.get("skip_merged_rows") and row["merged"]:
            continue
        out.append(tuple(None if is_blank(x) else x for x in row["values"][a:b + 1]))
    return out, (b - a + 1)


# =====================================================================
# PASTE INTO MAIN REPORT (values only)
# =====================================================================

def paste(report_wb, job, data, ncols):
    ws = report_wb.Worksheets(job["target_sheet"])
    tgt = ws.Range(job["target_cell"])
    r0, c0 = tgt.Row, tgt.Column

    if job.get("clear_old_data", True):
        ur = ws.UsedRange
        last = ur.Row + ur.Rows.Count - 1
        if last >= r0:
            ws.Range(ws.Cells(r0, c0), ws.Cells(last, c0 + ncols - 1)).ClearContents()

    if data:
        ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(data) - 1, c0 + ncols - 1)).Value = tuple(data)


# =====================================================================
# MAIN
# =====================================================================

def main():
    folder = base_folder()
    log(f"Folder: {folder}")
    log("")

    errors = []
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    report_wb = None

    try:
        report_path = find_file(folder, MAIN_REPORT_NAME)
        log(f"Main report: {os.path.basename(report_path)}")

        if MAKE_BACKUP:
            log(f"Backup saved: {backup_report(report_path)}")

        report_wb = excel.Workbooks.Open(report_path, 0, False)
        if report_wb.ReadOnly:
            raise RuntimeError("The main report is open somewhere else. Close it and run again.")

        for job in JOBS:
            log("")
            log(f"--- {job['name']} ---")
            try:
                website_path = find_file(folder, job["website_file"])
                log(f"Reading: {os.path.basename(website_path)}")

                wb = excel.Workbooks.Open(website_path, 0, True)
                try:
                    ws = wb.Worksheets(job["website_sheet"]) if job.get("website_sheet") else wb.Worksheets(1)
                    rows, last_col = read_sheet(ws)
                    data, ncols = transform(rows, last_col, job)
                finally:
                    wb.Close(SaveChanges=False)

                paste(report_wb, job, data, ncols)
                log(f"Pasted {len(data)} rows into '{job['target_sheet']}' at {job['target_cell']}")
            except Exception as e:
                errors.append(f"{job['name']}: {e}")
                log(f"ERROR: {e}")

        excel.CalculateFull()

        try:
            first = report_wb.Worksheets(1)
            first.Activate()
            first.Range("A1").Select()
        except Exception:
            pass

        report_wb.Save()
        log("")
        log("Report saved.")

    except Exception as e:
        errors.append(str(e))
        log(f"ERROR: {e}")
        log(traceback.format_exc())

    finally:
        try:
            if report_wb is not None:
                report_wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.ScreenUpdating = True
        excel.Quit()

    log("")
    if errors:
        log("Finished with errors:")
        for e in errors:
            log(f"  - {e}")
    else:
        log("All done, no errors.")

    try:
        with open(os.path.join(folder, "update_log.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(LOG_LINES))
    except Exception:
        pass


if __name__ == "__main__":
    main()
    input("\nPress Enter to close...")
