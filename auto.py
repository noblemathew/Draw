import os
import sys
import shutil
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox

import pythoncom
import win32com.client as win32


MAIN_REPORT_NAME = ""      
website_FOLDER_NAME = "website"                 
MATCH_CHARS = 10                          
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb")
MAKE_BACKUP = True                        
SAVE_website_CHANGES = True                  

# =====================================================================
# HELPERS
# =====================================================================

XL_TO_LEFT = -4159            


def base_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def website_folder():
    return os.path.join(base_folder(), website_FOLDER_NAME)


def col_num(letter):
    n = 0
    for ch in letter.upper().strip():
        n = n * 26 + (ord(ch) - 64)
    return n


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def excel_files(folder):
    if not os.path.isdir(folder):
        return []
    out = []
    for f in os.listdir(folder):
        if f.startswith("~$"):              
            continue
        if os.path.splitext(f)[1].lower() in EXCEL_EXTS:
            out.append(os.path.join(folder, f))
    return out


def find_main_report():
    key = MAIN_REPORT_NAME.lower()
    matches = [p for p in excel_files(base_folder())
               if os.path.splitext(os.path.basename(p))[0].lower().startswith(key)]
    if not matches:
        return None
    return max(matches, key=os.path.getmtime)


def match_key(name):
    return name[:MATCH_CHARS].strip().lower()


def find_website(name):
    key = match_key(name)
    matches = [p for p in excel_files(website_folder())
               if match_key(os.path.splitext(os.path.basename(p))[0]).startswith(key)]
    if not matches:
        return None
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


def key_col_of(row, cfg):
    """0-based column of the key if it is the FIRST value in the row, else None."""
    first = next((k for k, x in enumerate(row) if not is_blank(x)), None)
    if first is None:
        return None
    if str(row[first]).strip().upper() == str(cfg["key"]).strip().upper():
        return first
    return None


def fill_missing_labels(ws, job, log):
    cfg = job.get("fill_labels")
    if not cfg:
        return

    rows, _, _ = read_values(ws)
    start = job.get("align_from_row", 1)
    labels = cfg["labels"]
    n = len(labels)

    def val(r, c):                                       
        if r >= len(rows) or c >= len(rows[r]):
            return None
        return rows[r][c]

    filled = 0
    for r in range(start - 1, len(rows)):
        key_c = key_col_of(rows[r], cfg)
        if not key_c:                                     
            continue
        chk = key_c - 1                                   
        if r + n >= len(rows):                           
            continue
        if all(is_blank(val(r + 1 + i, chk)) for i in range(n)):
            for i, label in enumerate(labels):
                excel_row = r + 2 + i                     
                ws.Range(ws.Cells(excel_row, chk + 1), ws.Cells(excel_row, chk + 1)).Value = label
                rows[r + 1 + i][chk] = label
            filled += 1
    log(f"  '{cfg['key']}' label sets filled: {filled}")


def clean_website_sheet(ws, job, log):
    start = job.get("align_from_row", 1)

    ws.Cells.UnMerge()

    if job.get("remove_blank_rows", True):
        rows, last_row, _ = read_values(ws)
        removed = 0
        for r in range(last_row, start - 1, -1):         
            if all(is_blank(x) for x in rows[r - 1]):
                ws.Rows(r).Delete()
                removed += 1
        log(f"  Blank rows removed: {removed}")

    fill_missing_labels(ws, job, log)

    rows, last_row, _ = read_values(ws)

    cfg = job.get("fill_labels")

    moved = 0
    for r in range(start, last_row + 1):
        v = rows[r - 1]
        first = next((k for k, x in enumerate(v) if not is_blank(x)), None)
        if cfg and first and key_col_of(v, cfg) == first:
            first -= 1                                    
        if first:                                        
            ws.Range(ws.Cells(r, 1), ws.Cells(r, first)).Delete(XL_TO_LEFT)
            moved += 1
    log(f"  Rows moved to start at column A: {moved}")

    if job.get("remove_blank_columns", True):
        rows, last_row, last_col = read_values(ws)
        removed = 0
        if last_row >= start:
            for c in range(last_col, 0, -1):             
                if all(is_blank(rows[r - 1][c - 1]) for r in range(start, last_row + 1)):
                    ws.Range(ws.Cells(start, c), ws.Cells(last_row, c)).Delete(XL_TO_LEFT)
                    removed += 1
        log(f"  Blank columns removed: {removed}")

def get_copy_data(ws, job):
    rows, _, _ = read_values(ws)
    cp = job["copy"]
    a = col_num(cp["cols"][0]) - 1
    b = col_num(cp["cols"][1]) - 1

    out = []
    for i in range(cp["start_row"] - 1, len(rows)):
        row = rows[i] + [None] * (b + 1 - len(rows[i]))
        if all(is_blank(x) for x in row):                 
            break
        out.append(tuple(None if is_blank(x) else x for x in row[a:b + 1]))
    return out, (b - a + 1)


def get_website_data(excel, job, log, read_only):
    """Open the website file, clean it, return the data to paste."""
    path = find_website(job["website_file"])
    if not path:
        raise FileNotFoundError(f"No file starting with '{job['website_file'][:MATCH_CHARS]}' in the website folder")
    log(f"  website file: {os.path.basename(path)}")

    wb = excel.Workbooks.Open(path, 0, read_only)
    try:
        if not read_only and SAVE_website_CHANGES and wb.ReadOnly:
            raise RuntimeError("website file is open somewhere else. Close it and try again.")
        ws = wb.Worksheets(job["website_sheet"]) if job.get("website_sheet") else wb.Worksheets(1)

        clean_website_sheet(ws, job, log)
        if not read_only and SAVE_website_CHANGES:
            wb.Save()
            log("  website file cleaned and saved")

        return get_copy_data(ws, job)
    finally:
        wb.Close(SaveChanges=False)

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



def norm(v):
    if is_blank(v):
        return None
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return round(float(v), 6)
    s = str(v).strip()
    try:
        return round(float(s.replace(",", "")), 6)
    except ValueError:
        return s


def compare(report_wb, job, data, ncols, log):
    """Return number of differences. Also checks old leftover rows below the data."""
    ws = report_wb.Worksheets(job["target_sheet"])
    tgt = ws.Range(job["target_cell"])
    r0, c0 = tgt.Row, tgt.Column

    ur = ws.UsedRange
    last = max(ur.Row + ur.Rows.Count - 1, r0 + len(data) - 1, r0)
    vals = ws.Range(ws.Cells(r0, c0), ws.Cells(last, c0 + ncols - 1)).Value
    if not isinstance(vals, tuple):
        vals = ((vals,),)

    diffs = []
    for i, row in enumerate(vals):
        for j in range(ncols):
            expected = data[i][j] if i < len(data) else None
            if norm(expected) != norm(row[j]):
                diffs.append((r0 + i, c0 + j, expected, row[j]))

    for (r, c, e, a) in diffs[:5]:
        cell = ws.Cells(r, c).Address.replace("$", "")
        log(f"    {cell}: website = {e!r} | report = {a!r}")
    if len(diffs) > 5:
        log(f"    ... and {len(diffs) - 5} more")
    return len(diffs)



def new_excel():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    return excel


def run_update(log, status, progress):
    errors = []
    excel = new_excel()
    report_wb = None
    try:
        report_path = find_main_report()
        if not report_path:
            raise FileNotFoundError(f"Main report '{MAIN_REPORT_NAME}' not found in {base_folder()}")
        log(f"Main report: {os.path.basename(report_path)}")

        if MAKE_BACKUP:
            log(f"Backup saved: {os.path.basename(backup_report(report_path))}")

        report_wb = excel.Workbooks.Open(report_path, 0, False)
        if report_wb.ReadOnly:
            raise RuntimeError("The main report is open somewhere else. Close it and run again.")

        for i, job in enumerate(JOBS):
            log("")
            log(f"--- {job['name']} ---")
            status(i, "Running...", "busy")
            try:
                data, ncols = get_website_data(excel, job, log, read_only=False)
                paste(report_wb, job, data, ncols)
                log(f"  Pasted {len(data)} rows into '{job['target_sheet']}' at {job['target_cell']}")
                status(i, f"Done · {len(data)} rows", "ok")
            except Exception as e:
                errors.append(f"{job['name']}: {e}")
                log(f"  ERROR: {e}")
                status(i, "Error", "bad")
            progress(i + 1)

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
    return errors


def run_qc(log, status, progress):
    errors = []
    excel = new_excel()
    report_wb = None
    try:
        report_path = find_main_report()
        if not report_path:
            raise FileNotFoundError(f"Main report '{MAIN_REPORT_NAME}' not found in {base_folder()}")
        log(f"QC against: {os.path.basename(report_path)} (read only)")

        report_wb = excel.Workbooks.Open(report_path, 0, True)

        for i, job in enumerate(JOBS):
            log("")
            log(f"--- QC {job['name']} ---")
            status(i, "Checking...", "busy")
            try:
                data, ncols = get_website_data(excel, job, log, read_only=True)
                n = compare(report_wb, job, data, ncols, log)
                if n == 0:
                    log(f"  OK · {len(data)} rows match")
                    status(i, f"QC OK · {len(data)} rows", "ok")
                else:
                    log(f"  {n} cell(s) don't match")
                    errors.append(f"{job['name']}: {n} cell(s) don't match")
                    status(i, f"QC · {n} diffs", "bad")
            except Exception as e:
                errors.append(f"{job['name']}: {e}")
                log(f"  ERROR: {e}")
                status(i, "Error", "bad")
            progress(i + 1)

    except Exception as e:
        errors.append(str(e))
        log(f"ERROR: {e}")

    finally:
        try:
            if report_wb is not None:
                report_wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.ScreenUpdating = True
        excel.Quit()
    return errors


# =====================================================================
# UI
# =====================================================================

C = {
    "bg": "#F4F6FA",
    "card": "#FFFFFF",
    "border": "#E3E7EF",
    "text": "#1F2937",
    "muted": "#6B7280",
    "header": "#111827",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "qc": "#0F766E",
    "qc_hover": "#0B5F59",
    "ok": "#15803D",
    "bad": "#DC2626",
    "busy": "#D97706",
    "idle": "#6B7280",
    "log_bg": "#0F172A",
    "log_fg": "#CBD5E1",
}
FONT = "Segoe UI"


class FlatButton(tk.Label):

    def __init__(self, parent, text, command, bg, hover, fg="white", **kw):
        super().__init__(parent, text=text, bg=bg, fg=fg, cursor="hand2",
                         font=(FONT, 10, "bold"), padx=18, pady=8, **kw)
        self._bg, self._hover, self._cmd, self._enabled = bg, hover, command, True
        self.bind("<Enter>", lambda e: self._enabled and self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._bg if self._enabled else "#9CA3AF"))
        self.bind("<Button-1>", lambda e: self._enabled and self._cmd())

    def set_enabled(self, on):
        self._enabled = on
        self.config(bg=self._bg if on else "#9CA3AF", cursor="hand2" if on else "arrow")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("State Rx Report Updater")
        self.configure(bg=C["bg"])
        self.geometry("880x620")
        self.minsize(780, 540)
        self.q = queue.Queue()
        self.busy = False
        self.log_lines = []

        self._style()
        self._build()
        self.refresh()
        self.after(100, self._poll)

        # center on screen
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 3
        self.geometry(f"+{x}+{y}")

    # ---------- styles ----------
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Treeview", background=C["card"], fieldbackground=C["card"],
                    foreground=C["text"], rowheight=30, borderwidth=0, font=(FONT, 10))
        s.configure("Treeview.Heading", background="#F9FAFB", foreground=C["muted"],
                    font=(FONT, 9, "bold"), borderwidth=0, relief="flat", padding=(8, 6))
        s.map("Treeview.Heading", background=[("active", "#F3F4F6")])
        s.map("Treeview", background=[("selected", "#EEF2FF")], foreground=[("selected", C["text"])])
        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        s.configure("Bar.Horizontal.TProgressbar", troughcolor=C["border"], background=C["accent"],
                    bordercolor=C["border"], lightcolor=C["accent"], darkcolor=C["accent"], thickness=6)
        s.configure("Vertical.TScrollbar", background=C["border"], troughcolor=C["card"],
                    bordercolor=C["card"], arrowcolor=C["muted"])

    def _card(self, parent):
        return tk.Frame(parent, bg=C["card"], highlightbackground=C["border"], highlightthickness=1)

    # ---------- layout ----------
    def _build(self):
        # Header
        head = tk.Frame(self, bg=C["header"])
        head.pack(fill="x")
        tk.Label(head, text="State Rx Report Updater", bg=C["header"], fg="white",
                 font=(FONT, 15, "bold")).pack(side="left", padx=20, pady=(14, 14))
        tk.Label(head, text="website  →  report", bg=C["header"], fg="#9CA3AF",
                 font=(FONT, 10)).pack(side="left", pady=(18, 14))

        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=18, pady=14)

        # Report + folder card
        info = self._card(body)
        info.pack(fill="x")
        info.columnconfigure(1, weight=1)

        tk.Label(info, text="MAIN REPORT", bg=C["card"], fg=C["muted"],
                 font=(FONT, 8, "bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 0))
        self.report_lbl = tk.Label(info, text="", bg=C["card"], fg=C["text"], font=(FONT, 11, "bold"), anchor="w")
        self.report_lbl.grid(row=1, column=0, columnspan=2, sticky="w", padx=14)

        tk.Label(info, text="website FOLDER", bg=C["card"], fg=C["muted"],
                 font=(FONT, 8, "bold")).grid(row=2, column=0, sticky="w", padx=14, pady=(8, 0))
        self.folder_lbl = tk.Label(info, text="", bg=C["card"], fg=C["text"], font=(FONT, 9), anchor="w")
        self.folder_lbl.grid(row=3, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 10))

        tk.Button(info, text="⟳  Refresh", command=self.refresh, relief="flat", bd=0,
                  bg="#EEF2FF", fg=C["accent"], activebackground="#E0E7FF", cursor="hand2",
                  font=(FONT, 9, "bold"), padx=12, pady=4).grid(row=0, column=2, rowspan=2, padx=14, pady=10, sticky="e")
        tk.Button(info, text="📂  Open folder", command=self.open_folder, relief="flat", bd=0,
                  bg="#F3F4F6", fg=C["text"], activebackground="#E5E7EB", cursor="hand2",
                  font=(FONT, 9), padx=12, pady=4).grid(row=2, column=2, rowspan=2, padx=14, pady=(0, 10), sticky="e")

        # Files table
        table_card = self._card(body)
        table_card.pack(fill="both", expand=True, pady=(12, 0))

        cols = ("no", "file", "sheet", "status")
        self.tree = ttk.Treeview(table_card, columns=cols, show="headings", height=7, selectmode="none")
        for cid, title, w, anchor in [("no", "#", 40, "center"), ("file", "website FILE", 360, "w"),
                                       ("sheet", "TARGET SHEET", 160, "w"), ("status", "STATUS", 170, "w")]:
            self.tree.heading(cid, text=title, anchor=anchor)
            self.tree.column(cid, width=w, anchor=anchor, stretch=(cid == "file"))
        for tag in ("ok", "bad", "busy", "idle"):
            self.tree.tag_configure(tag, foreground=C[tag])
        sb = ttk.Scrollbar(table_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        sb.pack(side="right", fill="y")

        # Log
        log_card = tk.Frame(body, bg=C["log_bg"])
        log_card.pack(fill="both", pady=(12, 0))
        self.log_txt = tk.Text(log_card, height=7, bg=C["log_bg"], fg=C["log_fg"], bd=0,
                               font=("Consolas", 9), insertbackground=C["log_fg"], wrap="word",
                               padx=12, pady=8, state="disabled")
        self.log_txt.pack(fill="both", expand=True)
        self.log_txt.tag_configure("err", foreground="#F87171")
        self.log_txt.tag_configure("ok", foreground="#4ADE80")
        self.log_txt.tag_configure("head", foreground="#93C5FD")

        # Bottom bar
        bar = tk.Frame(self, bg=C["bg"])
        bar.pack(fill="x", padx=18, pady=(0, 14))
        self.status_lbl = tk.Label(bar, text="Ready", bg=C["bg"], fg=C["muted"], font=(FONT, 9))
        self.status_lbl.pack(side="left")
        self.run_btn = FlatButton(bar, "▶  Run", self.start_run, C["accent"], C["accent_hover"])
        self.run_btn.pack(side="right")
        self.qc_btn = FlatButton(bar, "✓  QC", self.start_qc, C["qc"], C["qc_hover"])
        self.qc_btn.pack(side="right", padx=(0, 10))
        self.pbar = ttk.Progressbar(bar, style="Bar.Horizontal.TProgressbar", length=180, mode="determinate")
        self.pbar.pack(side="right", padx=16)

    # ---------- data ----------
    def refresh(self):
        if self.busy:
            return
        rp = find_main_report()
        if rp:
            self.report_lbl.config(text=f"●  {os.path.basename(rp)}", fg=C["text"])
        else:
            self.report_lbl.config(text=f"●  '{MAIN_REPORT_NAME}' not found next to the app", fg=C["bad"])

        wf = website_folder()
        self.folder_lbl.config(text=wf if os.path.isdir(wf) else f"{wf}   (folder not found)",
                               fg=C["muted"] if os.path.isdir(wf) else C["bad"])

        self.tree.delete(*self.tree.get_children())
        found = 0
        for i, job in enumerate(JOBS):
            p = find_website(job["website_file"])
            if p:
                found += 1
                fname, st, tag = os.path.basename(p), "Ready", "idle"
            else:
                fname, st, tag = f"{job['website_file']}…   (not found)", "Missing", "bad"
            self.tree.insert("", "end", iid=str(i), values=(i + 1, fname, job["target_sheet"], f"●  {st}"),
                             tags=(tag,))
        self.status_lbl.config(text=f"{found} of {len(JOBS)} website files found")
        self.pbar.config(maximum=max(len(JOBS), 1), value=0)

    def open_folder(self):
        try:
            os.startfile(base_folder())
        except Exception:
            pass

    # ---------- thread-safe callbacks ----------
    def _log(self, msg=""):
        self.q.put(("log", msg))

    def _status(self, i, text, tag):
        self.q.put(("status", (i, text, tag)))

    def _progress(self, n):
        self.q.put(("progress", n))

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self.log_lines.append(val)
                    tag = ("err" if "ERROR" in val or "don't match" in val else
                           "ok" if val.strip().startswith("OK") or "Report saved" in val or "no errors" in val else
                           "head" if val.startswith("---") else None)
                    self.log_txt.config(state="normal")
                    self.log_txt.insert("end", val + "\n", tag)
                    self.log_txt.see("end")
                    self.log_txt.config(state="disabled")
                elif kind == "status":
                    i, text, tag = val
                    if self.tree.exists(str(i)):
                        v = list(self.tree.item(str(i), "values"))
                        v[3] = f"●  {text}"
                        self.tree.item(str(i), values=v, tags=(tag,))
                elif kind == "progress":
                    self.pbar.config(value=val)
                elif kind == "done":
                    self._finish(*val)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    # ---------- run / qc ----------
    def _start(self, worker, title):
        if self.busy:
            return
        if not find_main_report():
            messagebox.showerror("Main report missing", f"'{MAIN_REPORT_NAME}' was not found next to the app.")
            return
        self.refresh()
        self.busy = True
        self.run_btn.set_enabled(False)
        self.qc_btn.set_enabled(False)
        self.status_lbl.config(text=f"{title} in progress…", fg=C["busy"])
        self.log_lines = []
        self.log_txt.config(state="normal")
        self.log_txt.delete("1.0", "end")
        self.log_txt.config(state="disabled")

        def task():
            pythoncom.CoInitialize()
            try:
                self._log(f"{title} started {datetime.now():%d-%m-%Y %H:%M}")
                errors = worker(self._log, self._status, self._progress)
            except Exception as e:
                errors = [str(e)]
                self._log(f"ERROR: {e}")
            finally:
                pythoncom.CoUninitialize()
            self._log("")
            if errors:
                self._log(f"Finished with {len(errors)} problem(s).")
            else:
                self._log("All done, no errors.")
            self.q.put(("done", (title, errors)))

        threading.Thread(target=task, daemon=True).start()

    def start_run(self):
        if messagebox.askyesno("Run update",
                               "Clean the website files and update the State Rx report?\n\n"
                               "A backup of the report is saved first."):
            self._start(run_update, "Run")

    def start_qc(self):
        self._start(run_qc, "QC")

    def _finish(self, title, errors):
        self.busy = False
        self.run_btn.set_enabled(True)
        self.qc_btn.set_enabled(True)
        try:
            name = "update_log.txt" if title == "Run" else "qc_log.txt"
            with open(os.path.join(base_folder(), name), "w", encoding="utf-8") as f:
                f.write("\n".join(self.log_lines))
        except Exception:
            pass
        if errors:
            self.status_lbl.config(text=f"{title} finished with {len(errors)} problem(s)", fg=C["bad"])
            messagebox.showwarning(f"{title} finished", "\n".join(errors[:10]))
        else:
            self.status_lbl.config(text=f"{title} finished · all good", fg=C["ok"])


if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)     # sharp text on high-DPI screens
    except Exception:
        pass
    App().mainloop()
