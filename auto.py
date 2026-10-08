

XL_TO_LEFT = -4159           


def base_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def Website_folder():
    return os.path.join(base_folder(), Website_FOLDER_NAME)


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


def find_Website(name):
    key = match_key(name)
    matches = [p for p in excel_files(Website_folder())
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



def clean_Website_sheet(ws, job, log):
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
            for c in range(last_col, 0, -1):              # right to left
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


def get_Website_data(excel, job, log, read_only):
    """Open the Website file, clean it, return the data to paste."""
    path = find_Website(job["Website_file"])
    if not path:
        raise FileNotFoundError(f"No file starting with '{job['Website_file'][:MATCH_CHARS]}' in the Website folder")
    log(f"  Website file: {os.path.basename(path)}")

    wb = excel.Workbooks.Open(path, 0, read_only)
    try:
        if not read_only and SAVE_Website_CHANGES and wb.ReadOnly:
            raise RuntimeError("Website file is open somewhere else. Close it and try again.")
        ws = wb.Worksheets(job["Website_sheet"]) if job.get("Website_sheet") else wb.Worksheets(1)

        clean_Website_sheet(ws, job, log)
        if not read_only and SAVE_Website_CHANGES:
            wb.Save()
            log("  Website file cleaned and saved")

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
        log(f"    {cell}: Website = {e!r} | report = {a!r}")
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
                data, ncols = get_Website_data(excel, job, log, read_only=False)
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
                data, ncols = get_Website_data(excel, job, log, read_only=True)
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

import time

C = {
    "bg": "#F3F5F9",          
    "panel": "#FFFFFF",      
    "panel2": "#F1F4F9",     
    "row_alt": "#F8FAFC",
    "border": "#E2E7EF",
    "text": "#1E293B",
    "muted": "#64748B",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "qc": "#0D9488",
    "qc_hover": "#0F766E",
    "disabled": "#CBD5E1",
    "ok": "#059669",
    "bad": "#DC2626",
    "busy": "#D97706",
    "idle": "#64748B",
    "head": "#2563EB",
    "sb": "#D5DCE6",
    "sb_active": "#B8C2D1",
    "term_bg": "#FBFCFE",
    "term_fg": "#334155",
    "term_dim": "#A0AEC0",
}
FONT = "Segoe UI"
MONO = "Consolas"


def fmt_clock(sec):
    sec = int(sec)
    m, s = divmod(sec, 60)
    h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def fmt_took(sec):
    return f"{sec:.1f}s" if sec < 60 else f"{int(sec // 60)}m {int(sec % 60):02d}s"


def short_path(p, n=64):
    return p if len(p) <= n else p[:22] + " … " + p[-(n - 25):]


def round_rect(c, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return c.create_polygon(pts, smooth=True, **kw)


class RoundButton(tk.Canvas):
    """Rounded flat button."""

    def __init__(self, parent, text, command, color, hover, width=104, height=34):
        super().__init__(parent, width=width, height=height, bg=parent["bg"],
                         highlightthickness=0, bd=0, cursor="hand2")
        self.color, self.hover, self.cmd, self.enabled = color, hover, command, True
        self.shape = round_rect(self, 1, 1, width - 1, height - 1, 9, fill=color, outline="")
        self.create_text(width / 2, height / 2, text=text, fill="white", font=(FONT, 10, "bold"))
        self.bind("<Enter>", lambda e: self.enabled and self.itemconfig(self.shape, fill=self.hover))
        self.bind("<Leave>", lambda e: self.itemconfig(self.shape, fill=self.color if self.enabled else C["disabled"]))
        self.bind("<Button-1>", lambda e: self.enabled and self.cmd())

    def set_enabled(self, on):
        self.enabled = on
        self.itemconfig(self.shape, fill=self.color if on else C["disabled"])
        self.config(cursor="hand2" if on else "arrow")


class GhostButton(tk.Label):
    """Small text button for toolbars."""

    def __init__(self, parent, text, command):
        super().__init__(parent, text=text, bg=parent["bg"], fg=C["muted"], cursor="hand2",
                         font=(FONT, 9), padx=8, pady=3)
        self.bind("<Enter>", lambda e: self.config(fg=C["text"], bg=C["panel2"]))
        self.bind("<Leave>", lambda e: self.config(fg=C["muted"], bg=parent["bg"]))
        self.bind("<Button-1>", lambda e: command())


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("State Rx Updater")
        self.configure(bg=C["bg"])
        self.geometry("760x680")
        self.minsize(680, 540)

        self.q = queue.Queue()
        self.busy = False
        self.log_lines = []
        self.t0 = None
        self.job_t0 = {}

        self._style()
        self._build()
        self.refresh()
        self.after(100, self._poll)

        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 3
        self.geometry(f"+{x}+{y}")

    # ---------- styles ----------
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Treeview", background=C["panel"], fieldbackground=C["panel"], foreground=C["text"],
                    rowheight=24, borderwidth=0, font=(FONT, 9))
        s.configure("Treeview.Heading", background=C["panel2"], foreground=C["muted"],
                    font=(FONT, 8, "bold"), borderwidth=0, relief="flat", padding=(8, 4))
        s.map("Treeview.Heading", background=[("active", C["panel2"])])
        s.map("Treeview", background=[("selected", C["panel2"])], foreground=[("selected", C["text"])])
        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        s.configure("Thin.Horizontal.TProgressbar", troughcolor=C["border"], background=C["accent"],
                    bordercolor=C["bg"], lightcolor=C["accent"], darkcolor=C["accent"], thickness=4)
        s.configure("Dark.Vertical.TScrollbar", background=C["sb"], troughcolor=C["panel"],
                    bordercolor=C["panel"], lightcolor=C["sb"], darkcolor=C["sb"],
                    arrowcolor=C["muted"], gripcount=0, relief="flat", arrowsize=10)
        s.map("Dark.Vertical.TScrollbar", background=[("active", C["sb_active"])])

    def _panel(self, parent, **pack):
        f = tk.Frame(parent, bg=C["panel"], highlightbackground=C["border"],
                     highlightcolor=C["border"], highlightthickness=1)
        f.pack(**pack)
        return f

    def _section(self, parent, title, right=None):
        row = tk.Frame(parent, bg=C["bg"])
        row.pack(fill="x", pady=(12, 5))
        tk.Label(row, text=title, bg=C["bg"], fg=C["muted"], font=(FONT, 8, "bold")).pack(side="left")
        return row

    # ---------- layout ----------
    def _build(self):
        root = tk.Frame(self, bg=C["bg"])
        root.pack(fill="both", expand=True, padx=16, pady=(14, 12))

        # Title row
        top = tk.Frame(root, bg=C["bg"])
        top.pack(fill="x")
        tk.Label(top, text="Updater", bg=C["bg"], fg=C["text"],
                 font=(FONT, 14, "bold")).pack(side="left")
        tk.Label(top, text="  Website → report", bg=C["bg"], fg=C["muted"],
                 font=(FONT, 9)).pack(side="left", pady=(5, 0))
        self.found_pill = tk.Label(top, text="", bg=C["panel"], fg=C["text"],
                                   font=(FONT, 8, "bold"), padx=10, pady=3)
        self.found_pill.pack(side="right")

        # Info card
        info = self._panel(root, fill="x", pady=(10, 0))
        info.columnconfigure(1, weight=1)
        tk.Label(info, text="REPORT", bg=C["panel"], fg=C["muted"], font=(FONT, 8, "bold"),
                 width=8, anchor="w").grid(row=0, column=0, sticky="w", padx=(12, 0), pady=(8, 2))
        self.report_lbl = tk.Label(info, text="", bg=C["panel"], fg=C["text"], font=(FONT, 9, "bold"), anchor="w")
        self.report_lbl.grid(row=0, column=1, sticky="w", pady=(8, 2))
        tk.Label(info, text="Website", bg=C["panel"], fg=C["muted"], font=(FONT, 8, "bold"),
                 width=8, anchor="w").grid(row=1, column=0, sticky="w", padx=(12, 0), pady=(2, 8))
        self.folder_lbl = tk.Label(info, text="", bg=C["panel"], fg=C["muted"], font=(FONT, 9), anchor="w")
        self.folder_lbl.grid(row=1, column=1, sticky="w", pady=(2, 8))
        tools = tk.Frame(info, bg=C["panel"])
        tools.grid(row=0, column=2, rowspan=2, padx=8)
        GhostButton(tools, "⟳  Refresh", self.refresh).pack(side="left")
        GhostButton(tools, "📂  Folder", self.open_folder).pack(side="left")

        # Website files (compact)
        self._section(root, "Website FILES")
        table = self._panel(root, fill="x")
        cols = ("no", "file", "sheet", "status")
        self.tree = ttk.Treeview(table, columns=cols, show="headings", height=5, selectmode="none")
        for cid, title, w, anchor in [("no", "#", 34, "center"), ("file", "FILE", 300, "w"),
                                       ("sheet", "TARGET SHEET", 130, "w"), ("status", "STATUS", 190, "w")]:
            self.tree.heading(cid, text=title, anchor=anchor)
            self.tree.column(cid, width=w, anchor=anchor, stretch=(cid == "file"))
        for tag in ("ok", "bad", "busy", "idle"):
            self.tree.tag_configure(tag, foreground=C[tag])
        self.tree.tag_configure("alt", background=C["row_alt"])
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview, style="Dark.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")

        # Bottom bar (packed first so it is always visible)
        self.pbar = ttk.Progressbar(root, style="Thin.Horizontal.TProgressbar", mode="determinate")

        bar = tk.Frame(root, bg=C["bg"])
        bar.pack(side="bottom", fill="x")
        self.pbar.pack(side="bottom", fill="x", pady=(10, 8))
        self.timer_lbl = tk.Label(bar, text="⏱ 00:00", bg=C["bg"], fg=C["text"], font=(MONO, 13, "bold"))
        self.timer_lbl.pack(side="left")
        self.status_lbl = tk.Label(bar, text="Ready", bg=C["bg"], fg=C["muted"], font=(FONT, 9))
        self.status_lbl.pack(side="left", padx=(12, 0), pady=(3, 0))
        self.run_btn = RoundButton(bar, "▶  Run", self.start_run, C["accent"], C["accent_hover"])
        self.run_btn.pack(side="right")
        self.qc_btn = RoundButton(bar, "✓  QC", self.start_qc, C["qc"], C["qc_hover"], width=88)
        self.qc_btn.pack(side="right", padx=(0, 8))

        # Output (big)
        sec = self._section(root, "OUTPUT")
        GhostButton(sec, "Clear", self.clear_log).pack(side="right")
        GhostButton(sec, "Open log", self.open_log).pack(side="right")

        term = tk.Frame(root, bg=C["term_bg"], highlightbackground=C["border"],
                        highlightcolor=C["border"], highlightthickness=1)
        term.pack(fill="both", expand=True)
        self.log_txt = tk.Text(term, bg=C["term_bg"], fg=C["term_fg"], bd=0, font=(MONO, 9),
                               insertbackground=C["term_fg"], wrap="word", padx=12, pady=10,
                               state="disabled", spacing1=1, spacing3=1, highlightthickness=0, height=10)
        lsb = ttk.Scrollbar(term, orient="vertical", command=self.log_txt.yview, style="Dark.Vertical.TScrollbar")
        self.log_txt.configure(yscrollcommand=lsb.set)
        self.log_txt.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        for tag, col in (("err", C["bad"]), ("ok", C["ok"]), ("head", C["head"]),
                         ("dim", C["term_dim"]), ("time", C["busy"])):
            self.log_txt.tag_configure(tag, foreground=col)

    # ---------- data ----------
    def refresh(self):
        if self.busy:
            return
        rp = find_main_report()
        if rp:
            self.report_lbl.config(text=os.path.basename(rp), fg=C["text"])
        else:
            self.report_lbl.config(text=f"'{MAIN_REPORT_NAME}' not found next to the app", fg=C["bad"])

        wf = Website_folder()
        ok = os.path.isdir(wf)
        self.folder_lbl.config(text=short_path(wf) if ok else f"{short_path(wf, 48)}  (not found)",
                               fg=C["muted"] if ok else C["bad"])

        self.tree.delete(*self.tree.get_children())
        found = 0
        for i, job in enumerate(JOBS):
            p = find_Website(job["Website_file"])
            if p:
                found += 1
                fname, st, tag = os.path.basename(p), "Ready", "idle"
            else:
                fname, st, tag = f"{job['Website_file']}…  (not found)", "Missing", "bad"
            tags = (tag, "alt") if i % 2 else (tag,)
            self.tree.insert("", "end", iid=str(i), values=(i + 1, fname, job["target_sheet"], f"●  {st}"),
                             tags=tags)

        all_ok = found == len(JOBS)
        self.found_pill.config(text=f"{found} / {len(JOBS)} files found",
                               fg=C["ok"] if all_ok else C["bad"])
        self.pbar.config(maximum=max(len(JOBS), 1), value=0)

    def open_folder(self):
        try:
            os.startfile(base_folder())
        except Exception:
            pass

    def open_log(self):
        for name in ("update_log.txt", "qc_log.txt"):
            p = os.path.join(base_folder(), name)
            if os.path.exists(p):
                try:
                    os.startfile(p)
                except Exception:
                    pass
                return

    def clear_log(self):
        if self.busy:
            return
        self.log_txt.config(state="normal")
        self.log_txt.delete("1.0", "end")
        self.log_txt.config(state="disabled")

    def _write(self, msg, tag=None):
        self.log_txt.config(state="normal")
        if msg:
            self.log_txt.insert("end", datetime.now().strftime("%H:%M:%S  "), "dim")
        self.log_txt.insert("end", msg + "\n", tag)
        self.log_txt.see("end")
        self.log_txt.config(state="disabled")

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
                    tag = ("err" if "ERROR" in val or "don't match" in val or "problem" in val else
                           "ok" if val.strip().startswith("OK") or "Report saved" in val or "no errors" in val else
                           "head" if val.startswith("---") else None)
                    self._write(val, tag)
                elif kind == "status":
                    i, text, tag = val
                    if tag == "busy":
                        self.job_t0[i] = time.time()
                    elif i in self.job_t0:
                        text += f"  ·  {fmt_took(time.time() - self.job_t0.pop(i))}"
                    if self.tree.exists(str(i)):
                        v = list(self.tree.item(str(i), "values"))
                        v[3] = f"●  {text}"
                        tags = (tag, "alt") if i % 2 else (tag,)
                        self.tree.item(str(i), values=v, tags=tags)
                        self.tree.see(str(i))
                elif kind == "progress":
                    self.pbar.config(value=val)
                elif kind == "done":
                    self._finish(*val)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _tick(self):
        if self.busy and self.t0:
            self.timer_lbl.config(text=f"⏱ {fmt_clock(time.time() - self.t0)}")
            self.after(500, self._tick)

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
        self.timer_lbl.config(fg=C["busy"])
        self.log_lines = []
        self.job_t0 = {}
        self.clear_log_force()
        self.t0 = time.time()
        self._tick()

        def task():
            pythoncom.CoInitialize()
            try:
                self._log(f"{title} started")
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

    def clear_log_force(self):
        self.log_txt.config(state="normal")
        self.log_txt.delete("1.0", "end")
        self.log_txt.config(state="disabled")

    def start_run(self):
        if messagebox.askyesno("Run update",
                               "Clean the Website files and update the report?\n\n"
                               "A backup of the report is saved first."):
            self._start(run_update, "Run")

    def start_qc(self):
        self._start(run_qc, "QC")

    def _finish(self, title, errors):
        took = time.time() - self.t0 if self.t0 else 0
        self.busy = False
        self.run_btn.set_enabled(True)
        self.qc_btn.set_enabled(True)
        self.timer_lbl.config(text=f"⏱ {fmt_clock(took)}", fg=C["text"])

        line = f"{title} took {fmt_took(took)}"
        self.log_lines.append(line)
        self._write(line, "time")

        try:
            name = "update_log.txt" if title == "Run" else "qc_log.txt"
            with open(os.path.join(base_folder(), name), "w", encoding="utf-8") as f:
                f.write("\n".join(self.log_lines))
        except Exception:
            pass

        if errors:
            self.status_lbl.config(text=f"{title} · {len(errors)} problem(s) · {fmt_took(took)}", fg=C["bad"])
            messagebox.showwarning(f"{title} finished", "\n".join(errors[:10]))
        else:
            self.status_lbl.config(text=f"{title} done · all good · {fmt_took(took)}", fg=C["ok"])


if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)     # sharp text on high-DPI screens
    except Exception:
        pass
    App().mainloop()
