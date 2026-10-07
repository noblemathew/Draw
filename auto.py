LOG_LINES = []

SAVE_WEBI_CHANGES = True      # True = the cleaned data is saved back into the webi file
XL_TO_LEFT = -4159            # Excel constant: shift cells left


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


def read_values(ws):
    ur = ws.UsedRange
    last_row = ur.Row + ur.Rows.Count - 1
    last_col = ur.Column + ur.Columns.Count - 1

    data = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col)).Value
    if not isinstance(data, tuple):          # single cell sheet
        data = ((data,),)
    return [list(r) for r in data], last_row, last_col


# =====================================================================
# STEP 1: FIX THE WEBI FILE ITSELF
# =====================================================================

def clean_webi_sheet(ws, job):
    start = job.get("align_from_row", 1)

    # Unmerge everything so cells can move freely
    ws.Cells.UnMerge()

    # 1. Remove blank rows (from start row to the end)
    if job.get("remove_blank_rows", True):
        rows, last_row, _ = read_values(ws)
        removed = 0
        for r in range(last_row, start - 1, -1):          # bottom-up so row numbers don't move
            if all(is_blank(x) for x in rows[r - 1]):
                ws.Rows(r).Delete()
                removed += 1
        log(f"  Blank rows removed: {removed}")

    # 2. Every row from start row: remove the empty cells before the data -> data starts in A
    rows, last_row, _ = read_values(ws)
    moved = 0
    for r in range(start, last_row + 1):
        v = rows[r - 1]
        first = next((k for k, x in enumerate(v) if not is_blank(x)), None)
        if first:                                         # None = empty row, 0 = already in A
            ws.Range(ws.Cells(r, 1), ws.Cells(r, first)).Delete(XL_TO_LEFT)
            moved += 1
    log(f"  Rows moved to start at column A: {moved}")

    # 3. Remove columns that are fully empty (from start row to the end)
    if job.get("remove_blank_columns", True):
        rows, last_row, last_col = read_values(ws)
        removed = 0
        if last_row >= start:
            for c in range(last_col, 0, -1):              # right to left
                if all(is_blank(rows[r - 1][c - 1]) for r in range(start, last_row + 1)):
                    ws.Range(ws.Cells(start, c), ws.Cells(last_row, c)).Delete(XL_TO_LEFT)
                    removed += 1
        log(f"  Blank columns removed: {removed}")


# =====================================================================
# STEP 2: COPY FROM THE CLEANED WEBI SHEET
# =====================================================================

def get_copy_data(ws, job):
    rows, _, _ = read_values(ws)
    cp = job["copy"]
    a = col_num(cp["cols"][0]) - 1
    b = col_num(cp["cols"][1]) - 1

    out = []
    for i in range(cp["start_row"] - 1, len(rows)):
        row = rows[i] + [None] * (b + 1 - len(rows[i]))
        if all(is_blank(x) for x in row):                 # stop at first empty row
            break
        out.append(tuple(None if is_blank(x) else x for x in row[a:b + 1]))
    return out, (b - a + 1)


# =====================================================================
# STEP 3: PASTE INTO MAIN REPORT (values only)
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
                webi_path = find_file(folder, job["webi_file"])
                log(f"Webi file: {os.path.basename(webi_path)}")

                wb = excel.Workbooks.Open(webi_path, 0, False)
                try:
                    if SAVE_WEBI_CHANGES and wb.ReadOnly:
                        raise RuntimeError("Webi file is open somewhere else. Close it and run again.")

                    ws = wb.Worksheets(job["webi_sheet"]) if job.get("webi_sheet") else wb.Worksheets(1)

                    clean_webi_sheet(ws, job)
                    if SAVE_WEBI_CHANGES:
                        wb.Save()
                        log("  Webi file cleaned and saved")

                    data, ncols = get_copy_data(ws, job)
                finally:
                    wb.Close(SaveChanges=False)

                paste(report_wb, job, data, ncols)
                log(f"  Pasted {len(data)} rows into '{job['target_sheet']}' at {job['target_cell']}")
            except Exception as e:
                errors.append(f"{job['name']}: {e}")
                log(f"  ERROR: {e}")

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
