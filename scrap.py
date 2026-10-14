XL_TO_LEFT = -4159


def base_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def logo_path():
    p = os.path.join(base_folder(), LOGO_FILE)
    if os.path.exists(p):
        return p
    return os.path.join(getattr(sys, "_MEIPASS", base_folder()), LOGO_FILE)


def website_folder():
    return os.path.join(base_folder(), website_FOLDER_NAME)


def col_num(letter):
    n = 0
    for ch in letter.upper().strip():
        n = n * 26 + (ord(ch) - 64)
    return n


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def is_number(v):
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        try:
            float(v.strip().replace(",", "").replace("%", ""))
            return True
        except ValueError:
            return False
    return False


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
    key = MAIN_REPORT_NAME.strip().lower()
    if not key:
        return None
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
    if not isinstance(data, tuple):
        data = ((data,),)
    return [list(r) for r in data], last_row, last_col


def set_cell(ws, r, c, value):
    ws.Range(ws.Cells(r, c), ws.Cells(r, c)).Value = value


def keep_keys(job):
    keys = list(job.get("keep_blank_before") or [])
    if job.get("fill_labels"):
        keys.append(job["fill_labels"]["key"])
    return {str(k).strip().upper() for k in keys}


def key_col_of(row, keys):
    first = next((k for k, x in enumerate(row) if not is_blank(x)), None)
    if first is None:
        return None
    if str(row[first]).strip().upper() in keys:
        return first
    return None


def fill_missing_labels(ws, job, log):
    cfg = job.get("fill_labels")
    if not cfg:
        return

    rows, _, _ = read_values(ws)
    start = job.get("align_from_row", 1)
    keys = {str(cfg["key"]).strip().upper()}
    labels = cfg["labels"]
    n = len(labels)

    def val(r, c):
        if r >= len(rows) or c >= len(rows[r]):
            return None
        return rows[r][c]

    filled = 0
    for r in range(start - 1, len(rows)):
        key_c = key_col_of(rows[r], keys)
        if not key_c:
            continue
        chk = key_c - 1
        if r + n >= len(rows):
            continue
        if all(is_blank(val(r + 1 + i, chk)) for i in range(n)):
            for i, label in enumerate(labels):
                set_cell(ws, r + 2 + i, chk + 1, label)
                rows[r + 1 + i][chk] = label
            filled += 1
    log(f"  '{cfg['key']}' label sets filled: {filled}")


def add_suffix(ws, job, log):
    cfg = job.get("add_suffix")
    if not cfg:
        return
    text = cfg["text"]
    rows, last_row, _ = read_values(ws)
    done = 0
    for r in cfg["rows"]:
        if r < 1 or r > last_row:
            continue
        for c, v in enumerate(rows[r - 1]):
            if isinstance(v, str) and v.strip() and not is_number(v) \
                    and not v.rstrip().endswith(text.strip()):
                set_cell(ws, r, c + 1, v.rstrip() + text)
                done += 1
    log(f"  '{text.strip()}' added to {done} cell(s)")


def is_total_row(row):
    return any(isinstance(x, str) and x.strip().lower() == "total" for x in row)


def is_zero(v):
    return is_number(v) and float(str(v).replace(",", "").replace("%", "")) == 0


def add_row_after_total(ws, job, log):
    mode = job.get("row_after_total")
    if mode not in ("blank", "zero"):
        return
    start = job.get("align_from_row", 1)
    rows, last_row, last_col = read_values(ws)

    added = 0
    for r in range(last_row, start - 1, -1):
        if not is_total_row(rows[r - 1]):
            continue
        nxt = rows[r] if r < last_row else None
        if mode == "blank":
            if nxt is not None and all(is_blank(x) for x in nxt):
                continue
            ws.Rows(r + 1).Insert()
        else:
            if nxt is not None and all(is_zero(x) for x in nxt[:last_col]):
                continue
            ws.Rows(r + 1).Insert()
            ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, last_col)).Value = 0
        added += 1
    log(f"  {'Blank' if mode == 'blank' else '0'} rows added below Total: {added}")


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
    keys = keep_keys(job)

    moved = 0
    for r in range(start, last_row + 1):
        v = rows[r - 1]
        first = next((k for k, x in enumerate(v) if not is_blank(x)), None)
        if keys and first and key_col_of(v, keys) == first:
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

    add_suffix(ws, job, log)
    add_row_after_total(ws, job, log)


def get_copy_data(ws, job):
    rows, _, _ = read_values(ws)
    cp = job["copy"]
    a = col_num(cp["cols"][0]) - 1
    b = col_num(cp["cols"][1]) - 1
    stop_at_blank = job.get("row_after_total") != "blank"

    picked = []
    for i in range(cp["start_row"] - 1, len(rows)):
        row = rows[i] + [None] * (b + 1 - len(rows[i]))
        part = tuple(None if is_blank(x) else x for x in row[a:b + 1])
        if all(is_blank(x) for x in row):
            if stop_at_blank:
                break
        picked.append(part)

    while picked and all(x is None for x in picked[-1]):
        picked.pop()
    return picked, (b - a + 1)


def get_website_data(excel, job, log, read_only):
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


def run_update(log, status, progress, stop):
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
            if stop():
                log("")
                log("Stopped by user. The report was NOT saved.")
                return errors
            log("")
            log(f"--- {job['name']} ---")
            status(i, "Updating")
            try:
                data, ncols = get_website_data(excel, job, log, read_only=False)
                paste(report_wb, job, data, ncols)
                log(f"  Pasted {len(data)} rows into '{job['target_sheet']}' at {job['target_cell']}")
            except Exception as e:
                errors.append(f"{job['name']}: {e}")
                log(f"  ERROR: {e}")
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


def run_qc(log, status, progress, stop):
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
            if stop():
                log("")
                log("QC stopped by user.")
                return errors
            log("")
            log(f"--- QC {job['name']} ---")
            status(i, "Checking")
            try:
                data, ncols = get_website_data(excel, job, log, read_only=True)
                n = compare(report_wb, job, data, ncols, log)
                if n == 0:
                    log(f"  OK · {len(data)} rows match")
                else:
                    log(f"  {n} cell(s) don't match")
                    errors.append(f"{job['name']}: {n} cell(s) don't match")
            except Exception as e:
                errors.append(f"{job['name']}: {e}")
                log(f"  ERROR: {e}")
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


C = {
    "bg": "#FFFFFF",
    "footer": "#F7F8FA",
    "line": "#E4E7EC",
    "field": "#D0D5DD",
    "text": "#101828",
    "sub": "#475467",
    "muted": "#98A2B3",
    "blue": "#1F4FB8",
    "blue_hover": "#183F94",
    "yellow": "#F6B820",
    "yellow_hover": "#E3A70F",
    "red": "#D92D20",
    "red_soft": "#FEF3F2",
    "disabled": "#F2F4F7",
    "disabled_text": "#98A2B3",
    "track": "#EAECF0",
    "ok": "#039855",
    "bad": "#D92D20",
    "busy": "#DC8B00",
    "win_hover": "#F2F4F7",
    "close_hover": "#E81123",
}
FONT = "Segoe UI"
UI_SCALE = 1.0


def px(n):
    return int(round(n * UI_SCALE))


def fmt_clock(sec):
    m, s = divmod(int(sec), 60)
    h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def fmt_took(sec):
    return f"{sec:.1f}s" if sec < 60 else f"{int(sec // 60)}m {int(sec % 60):02d}s"


def round_rect(c, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return c.create_polygon(pts, smooth=True, **kw)


def load_logo(w, h):
    path = logo_path()
    if not os.path.exists(path):
        return None
    try:
        from PIL import Image, ImageTk
        img = Image.open(path).convert("RGBA")
        ratio = min(w / img.width, h / img.height)
        img = img.resize((max(1, int(img.width * ratio)), max(1, int(img.height * ratio))), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        try:
            return tk.PhotoImage(file=path)
        except Exception:
            return None


class Button(tk.Canvas):

    def __init__(self, parent, text, command, style="primary", width=92, height=34):
        w, h = px(width), px(height)
        super().__init__(parent, width=w, height=h, bg=parent["bg"],
                         highlightthickness=0, bd=0, cursor="hand2")
        self.cmd, self.enabled = command, True
        s = {
            "primary": (C["blue"], C["blue_hover"], "#FFFFFF", ""),
            "accent": (C["yellow"], C["yellow_hover"], C["text"], ""),
            "outline": ("#FFFFFF", C["red_soft"], C["red"], C["field"]),
        }[style]
        self.fill, self.hover, self.fg, self.border = s
        self.shape = round_rect(self, 1, 1, w - 1, h - 1, px(6),
                                fill=self.fill, outline=self.border or self.fill)
        self.label = self.create_text(w / 2, h / 2, text=text, fill=self.fg, font=(FONT, 9, "bold"))
        self.bind("<Enter>", lambda e: self.enabled and self.itemconfig(self.shape, fill=self.hover))
        self.bind("<Leave>", lambda e: self.enabled and self.itemconfig(self.shape, fill=self.fill))
        self.bind("<Button-1>", lambda e: self.enabled and self.cmd())

    def set_enabled(self, on):
        self.enabled = on
        self.itemconfig(self.shape, fill=self.fill if on else C["disabled"],
                        outline=(self.border or self.fill) if on else C["line"])
        self.itemconfig(self.label, fill=self.fg if on else C["disabled_text"])
        self.config(cursor="hand2" if on else "arrow")


class WinButton(tk.Canvas):

    def __init__(self, parent, kind, command):
        w, h = px(44), px(30)
        super().__init__(parent, width=w, height=h, bg=C["bg"], highlightthickness=0, bd=0)
        self.kind = kind
        cx, cy, d = w / 2, h / 2, px(5)
        lw = max(1, px(1))
        if kind == "min":
            self.items = [self.create_line(cx - d, cy, cx + d, cy, fill=C["sub"], width=lw)]
            self.hover_bg, self.hover_fg = C["win_hover"], C["text"]
        else:
            self.items = [self.create_line(cx - d, cy - d, cx + d, cy + d, fill=C["sub"], width=lw),
                          self.create_line(cx - d, cy + d, cx + d, cy - d, fill=C["sub"], width=lw)]
            self.hover_bg, self.hover_fg = C["close_hover"], "#FFFFFF"
        self.bind("<Enter>", lambda e: self._paint(self.hover_bg, self.hover_fg))
        self.bind("<Leave>", lambda e: self._paint(C["bg"], C["sub"]))
        self.bind("<ButtonRelease-1>", lambda e: command())

    def _paint(self, bg, fg):
        self.config(bg=bg)
        for i in self.items:
            self.itemconfig(i, fill=fg)


class ThinBar(tk.Canvas):

    def __init__(self, parent, width, height=4):
        super().__init__(parent, width=width, height=px(height), bg=parent["bg"], highlightthickness=0, bd=0)
        self.w, self.h, self.maximum = width, px(height), 1
        r = self.h / 2
        round_rect(self, 0, 0, width, self.h, r, fill=C["track"], outline="")
        self.bar = round_rect(self, 0, 0, 0, self.h, r, fill=C["blue"], outline="")

    def set(self, value, maximum=None):
        if maximum:
            self.maximum = maximum
        x = max(0, min(1, value / max(self.maximum, 1))) * self.w
        self.coords(self.bar, *self._pts(x))
        self.itemconfig(self.bar, state="normal" if x > 0 else "hidden")

    def _pts(self, x2):
        h = self.h
        r = h / 2
        x2 = max(x2, 2 * r)
        return [r, 0, x2 - r, 0, x2, 0, x2, r, x2, h - r, x2, h, x2 - r, h, r, h, 0, h, 0, h - r, 0, r, 0, 0]


class LinkLabel(tk.Label):
    def __init__(self, parent, text, command):
        super().__init__(parent, text=text, bg=parent["bg"], fg=C["sub"], cursor="hand2", font=(FONT, 9))
        self.bind("<Enter>", lambda e: self.config(fg=C["blue"], font=(FONT, 9, "underline")))
        self.bind("<Leave>", lambda e: self.config(fg=C["sub"], font=(FONT, 9)))
        self.bind("<Button-1>", lambda e: command())


class App(tk.Tk):
    W, H = 640, 400
    PAD = 32

    def __init__(self):
        super().__init__()
        global UI_SCALE
        try:
            UI_SCALE = max(1.0, self.winfo_fpixels("1i") / 96.0)
        except Exception:
            UI_SCALE = 1.0

        self.title("Excel report Updater")
        self.configure(bg=C["line"])
        self.overrideredirect(True)
        w, h = px(self.W), px(self.H)
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.q = queue.Queue()
        self.busy = False
        self.stop_evt = threading.Event()
        self.log_lines = []
        self.t0 = None
        self.mode = ""
        self._minimized = False
        self._dx = self._dy = 0

        self._build()
        self.refresh()
        self.bind("<Map>", self._on_map)
        self.after(10, self._show_in_taskbar)
        self.after(100, self._poll)

    def _hwnd(self):
        import ctypes
        return ctypes.windll.user32.GetParent(self.winfo_id())

    def _show_in_taskbar(self):
        try:
            import ctypes
            user32 = ctypes.windll.user32
            hwnd = self._hwnd()
            ex = user32.GetWindowLongW(hwnd, -20)
            ex = (ex & ~0x00000080) | 0x00040000
            user32.SetWindowLongW(hwnd, -20, ex)
            self.withdraw()
            self.after(10, self.deiconify)
        except Exception:
            pass

    def minimize(self):
        self._minimized = True
        self.overrideredirect(False)
        self.iconify()

    def _on_map(self, event=None):
        if self._minimized and self.state() == "normal":
            self._minimized = False
            self.overrideredirect(True)
            self.after(10, self._show_in_taskbar)

    def _drag_start(self, e):
        self._dx, self._dy = e.x_root - self.winfo_x(), e.y_root - self.winfo_y()

    def _drag(self, e):
        self.geometry(f"+{e.x_root - self._dx}+{e.y_root - self._dy}")

    def _draggable(self, *widgets):
        for w in widgets:
            w.bind("<ButtonPress-1>", self._drag_start)
            w.bind("<B1-Motion>", self._drag)

    def on_close(self):
        if self.busy:
            messagebox.showinfo("Still working", "Please press Stop and wait until it finishes.", parent=self)
            return
        self.destroy()

    def _build(self):
        main = tk.Frame(self, bg=C["bg"])
        main.pack(fill="both", expand=True, padx=1, pady=1)
        W = px(self.W) - 2
        P = px(self.PAD)

        strip = tk.Canvas(main, width=W, height=px(3), bg=C["blue"], highlightthickness=0, bd=0)
        strip.pack(fill="x")
        strip.create_rectangle(W - px(120), 0, W, px(3), fill=C["yellow"], outline="")

        head = tk.Frame(main, bg=C["bg"])
        head.pack(fill="x")
        lw, lh = px(267), px(148)
        logo_box = tk.Frame(head, bg=C["bg"], width=lw, height=lh)
        logo_box.pack(pady=(px(6), 0))
        logo_box.pack_propagate(False)
        self.logo_img = load_logo(lw, lh)
        drag_items = [strip, head, logo_box]
        if self.logo_img is not None:
            lbl = tk.Label(logo_box, image=self.logo_img, bg=C["bg"], bd=0)
            lbl.pack(expand=True)
            drag_items.append(lbl)
        else:
            cv = tk.Canvas(logo_box, width=lw, height=lh, bg=C["bg"], highlightthickness=0)
            cv.pack()
            cv.create_rectangle(1, 1, lw - 1, lh - 1, outline=C["line"], dash=(4, 3))
            cv.create_text(lw / 2, lh * 0.45, text="LOGO", fill=C["muted"], font=(FONT, 14, "bold"))
            cv.create_text(lw / 2, lh * 0.6, text=f"{LOGO_FILE}  ·  267 × 148", fill=C["muted"], font=(FONT, 8))
            drag_items.append(cv)

        name = tk.Label(head, text="Excel report Updater", bg=C["bg"], fg=C["text"], font=(FONT, 13, "bold"))
        name.pack(pady=(px(4), 0))
        drag_items.append(name)
        self._draggable(*drag_items)

        controls = tk.Frame(main, bg=C["bg"])
        controls.place(relx=1.0, x=0, y=px(3), anchor="ne")
        WinButton(controls, "min", self.minimize).pack(side="left")
        WinButton(controls, "close", self.on_close).pack(side="left")

        footer = tk.Frame(main, bg=C["footer"], height=px(56))
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)
        tk.Frame(footer, bg=C["line"], height=1).pack(side="top", fill="x")
        fin = tk.Frame(footer, bg=C["footer"])
        fin.pack(fill="both", expand=True, padx=P)
        LinkLabel(fin, "View log", self.open_log).pack(side="left")
        tk.Label(fin, text="|", bg=C["footer"], fg=C["line"], font=(FONT, 9)).pack(side="left", padx=px(8))
        LinkLabel(fin, "Open folder", self.open_folder).pack(side="left")

        self.run_btn = Button(fin, "Run", self.start_run, "primary", width=96)
        self.run_btn.pack(side="right")
        self.qc_btn = Button(fin, "QC Check", self.start_qc, "accent", width=96)
        self.qc_btn.pack(side="right", padx=(0, px(8)))
        self.stop_btn = Button(fin, "Stop", self.stop, "outline", width=76)
        self.stop_btn.pack(side="right", padx=(0, px(8)))
        self.stop_btn.set_enabled(False)

        tk.Frame(main, bg=C["line"], height=1).pack(fill="x", padx=P, pady=(px(6), 0))
        body = tk.Frame(main, bg=C["bg"])
        body.pack(fill="both", expand=True, padx=P, pady=(px(10), 0))

        r1 = tk.Frame(body, bg=C["bg"])
        r1.pack(fill="x")
        tk.Label(r1, text="Report file", bg=C["bg"], fg=C["sub"], font=(FONT, 9), width=10,
                 anchor="w").pack(side="left")
        self.report_var = tk.StringVar(value=MAIN_REPORT_NAME)
        self.entry = tk.Entry(r1, textvariable=self.report_var, font=(FONT, 10),
                              bg="#FFFFFF", fg=C["text"], relief="flat", insertbackground=C["text"],
                              highlightthickness=1, highlightbackground=C["field"], highlightcolor=C["blue"],
                              disabledbackground=C["footer"], disabledforeground=C["muted"])
        self.entry.pack(side="left", fill="x", expand=True, ipady=px(4), padx=(0, px(12)))
        self.entry.bind("<Return>", lambda e: (self.focus_set(), self.refresh()))
        self.entry.bind("<FocusOut>", lambda e: self.refresh())
        self.report_ok = tk.Label(r1, text="", bg=C["bg"], font=(FONT, 9), width=10, anchor="w")
        self.report_ok.pack(side="left")

        r2 = tk.Frame(body, bg=C["bg"])
        r2.pack(fill="x", pady=(px(10), 0))
        d = px(8)
        self.dot = tk.Canvas(r2, width=d, height=d, bg=C["bg"], highlightthickness=0)
        self.dot_id = self.dot.create_oval(0, 0, d, d, fill=C["ok"], outline="")
        self.dot.pack(side="left", pady=(px(2), 0))
        self.status_lbl = tk.Label(r2, text="", bg=C["bg"], fg=C["text"], font=(FONT, 10, "bold"))
        self.status_lbl.pack(side="left", padx=(px(8), 0))
        self.timer_lbl = tk.Label(r2, text="00:00", bg=C["bg"], fg=C["sub"], font=(FONT, 10))
        self.timer_lbl.pack(side="right")
        tk.Label(r2, text="Elapsed", bg=C["bg"], fg=C["muted"], font=(FONT, 8)).pack(side="right", padx=(0, px(6)))

        self.pbar = ThinBar(body, width=W - 2 * P)
        self.pbar.pack(fill="x", pady=(px(8), 0))
        self.step_lbl = tk.Label(body, text="", bg=C["bg"], fg=C["sub"], font=(FONT, 9), anchor="w")
        self.step_lbl.pack(fill="x", pady=(px(6), 0))

        controls.lift()

    def set_status(self, text, color):
        self.status_lbl.config(text=text)
        self.dot.itemconfig(self.dot_id, fill=color)

    def refresh(self):
        if self.busy:
            return
        global MAIN_REPORT_NAME
        MAIN_REPORT_NAME = self.report_var.get().strip() or MAIN_REPORT_NAME
        rp = find_main_report()
        if rp:
            self.report_ok.config(text="✓  Found", fg=C["ok"])
        else:
            self.report_ok.config(text="✕  Not found", fg=C["bad"])

        missing = [j["name"] for j in JOBS if not find_website(j["website_file"])]
        found = len(JOBS) - len(missing)
        if rp and not missing:
            self.set_status("Ready", C["ok"])
        else:
            self.set_status("Action needed", C["bad"])
        self.step_lbl.config(
            text=f"{found} of {len(JOBS)} website files found" +
                 (f"  ·  missing: {', '.join(missing[:3])}{'…' if len(missing) > 3 else ''}" if missing else ""))
        self.pbar.set(0, len(JOBS))

    def _set_busy(self, on):
        self.busy = on
        self.run_btn.set_enabled(not on)
        self.qc_btn.set_enabled(not on)
        self.stop_btn.set_enabled(on)
        self.entry.config(state="disabled" if on else "normal")

    def open_folder(self):
        try:
            os.startfile(base_folder())
        except Exception:
            pass

    def open_log(self):
        names = ["update_log.txt", "qc_log.txt"]
        paths = [os.path.join(base_folder(), n) for n in names]
        paths = [p for p in paths if os.path.exists(p)]
        if not paths:
            messagebox.showinfo("No log yet", "Run or QC first.", parent=self)
            return
        try:
            os.startfile(max(paths, key=os.path.getmtime))
        except Exception:
            pass

    def _log(self, msg=""):
        self.q.put(("log", msg))

    def _status(self, i, text):
        self.q.put(("step", f"{text} {JOBS[i]['name']}  ·  {i + 1} of {len(JOBS)}"))

    def _progress(self, n):
        self.q.put(("progress", n))

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self.log_lines.append(val)
                elif kind == "step":
                    self.step_lbl.config(text=val)
                elif kind == "progress":
                    self.pbar.set(val)
                elif kind == "done":
                    self._finish(val)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _tick(self):
        if self.busy and self.t0:
            self.timer_lbl.config(text=fmt_clock(time.time() - self.t0))
            self.after(500, self._tick)

    def _start(self, worker, title):
        if self.busy:
            return
        self.refresh()
        if not find_main_report():
            messagebox.showerror("Report not found",
                                 f"No file starting with '{MAIN_REPORT_NAME}' was found next to the app.",
                                 parent=self)
            return

        self.mode = title
        self.stop_evt.clear()
        self.log_lines = [f"{title} started {datetime.now():%d-%m-%Y %H:%M:%S}"]
        self._set_busy(True)
        self.set_status("Updating report" if title == "Run" else "Running QC check", C["busy"])
        self.step_lbl.config(text="Opening Excel…")
        self.timer_lbl.config(fg=C["text"])
        self.pbar.set(0, len(JOBS))
        self.t0 = time.time()
        self._tick()

        def task():
            pythoncom.CoInitialize()
            try:
                errors = worker(self._log, self._status, self._progress, self.stop_evt.is_set)
            except Exception as e:
                errors = [str(e)]
                self._log(f"ERROR: {e}")
            finally:
                pythoncom.CoUninitialize()
            self.q.put(("done", errors))

        threading.Thread(target=task, daemon=True).start()

    def start_run(self):
        if messagebox.askyesno("Run", "Update the report now?\n\nA backup copy is saved first.", parent=self):
            self._start(run_update, "Run")

    def start_qc(self):
        self._start(run_qc, "QC")

    def stop(self):
        if self.busy and not self.stop_evt.is_set():
            self.stop_evt.set()
            self.stop_btn.set_enabled(False)
            self.set_status("Stopping", C["busy"])
            self.step_lbl.config(text="Finishing the current file, then stopping…")

    def _finish(self, errors):
        took = time.time() - self.t0 if self.t0 else 0
        title = self.mode
        stopped = self.stop_evt.is_set()
        self._set_busy(False)
        self.timer_lbl.config(text=fmt_clock(took), fg=C["sub"])

        self.log_lines.append(f"{title} took {fmt_took(took)}")
        try:
            name = "update_log.txt" if title == "Run" else "qc_log.txt"
            with open(os.path.join(base_folder(), name), "w", encoding="utf-8") as f:
                f.write("\n".join(self.log_lines))
        except Exception:
            pass

        if stopped:
            self.set_status("Stopped", C["busy"])
            self.step_lbl.config(text=f"Stopped after {fmt_took(took)}"
                                      + ("  ·  report not changed" if title == "Run" else ""))
        elif errors:
            self.set_status(f"{title} completed with {len(errors)} issue(s)", C["bad"])
            self.step_lbl.config(text=f"Took {fmt_took(took)}  ·  click View log for details")
            messagebox.showwarning(f"{title} completed", "\n".join(errors[:10]), parent=self)
        else:
            done = "Report updated successfully" if title == "Run" else "QC passed  ·  all data matches"
            self.set_status(done, C["ok"])
            self.pbar.set(len(JOBS))
            self.step_lbl.config(text=f"Completed in {fmt_took(took)}  ·  {len(JOBS)} of {len(JOBS)} files")


if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    App().mainloop()
