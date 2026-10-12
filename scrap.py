import sys
import time
import datetime
import decimal
from contextlib import contextmanager

import pyodbc
from PySide6.QtCore import Qt, Signal, QPointF, QRectF, QPoint
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap, QIcon, QFont, QBrush
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QStackedWidget, QFrame, QLabel, QLineEdit,
    QPushButton, QComboBox, QVBoxLayout, QHBoxLayout, QGridLayout, QListWidget,
    QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QAbstractItemView, QSizePolicy,
)


APP_TITLE = "My App"

USER_COLUMN = "UserName"      
CLASS_COLUMN = "Class"     
FLAG_COLUMN = "FlagColor"     

USER_LABEL, USER_PLACEHOLDER = "User", "Search users..."
CLASS_LABEL, CLASS_PLACEHOLDER = "Select Option for Class", "Search classes..."
FLAG_LABEL, FLAG_PLACEHOLDER = "Selection for Flag Color", "Search flag colors..."

RESULT_COLUMNS = [
    ("UserName", "User"),             
    ("Class", "Class"),               
    ("FlagColor", "Flag Color"),      
    ("Description", "Description"),   
    ("CreatedOn", "Created On"),      
]

MAX_RESULT_ROWS = 5000                
DATE_FORMAT = "%d-%m-%Y %H:%M"        


def pick_driver():
    preferred = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]
    installed = pyodbc.drivers()
    for d in preferred:
        if d in installed:
            return d
    raise RuntimeError(
        "No SQL Server ODBC driver found.\n"
        "Install 'ODBC Driver 18 for SQL Server' from Microsoft and try again."
    )


def db_connect(server, database="master"):
    driver = pick_driver()
    db = database.replace("}", "}}")
    cs = (
        f"DRIVER={{{driver}}};SERVER={server};DATABASE={{{db}}};"
        "Trusted_Connection=yes;"
    )
    if driver.startswith("ODBC Driver"):
        cs += "TrustServerCertificate=yes;"
    return pyodbc.connect(cs, timeout=10)


def q(name):
    return "[" + str(name).replace("]", "]]") + "]"


def fmt_value(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime(DATE_FORMAT)
    if isinstance(v, datetime.date):
        return v.strftime(DATE_FORMAT.split(" ")[0])
    if isinstance(v, decimal.Decimal):
        return format(v.normalize(), "f")
    return str(v)


@contextmanager
def busy():
    QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        yield
    finally:
        QApplication.restoreOverrideCursor()


def make_icon(kind, color="#4b5468", size=18):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    s = size
    pen = QPen(QColor(color))
    pen.setWidthF(max(1.6, s / 10))
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)

    if kind == "search":
        p.drawEllipse(QRectF(s * 0.12, s * 0.12, s * 0.56, s * 0.56))
        p.drawLine(QPointF(s * 0.62, s * 0.62), QPointF(s * 0.88, s * 0.88))
    elif kind == "chevron":
        p.drawPolyline([QPointF(s * 0.25, s * 0.38), QPointF(s * 0.5, s * 0.63),
                        QPointF(s * 0.75, s * 0.38)])
    elif kind == "refresh":
        p.drawArc(QRectF(s * 0.18, s * 0.18, s * 0.64, s * 0.64), 40 * 16, 280 * 16)
        x, y = s * 0.745, s * 0.294
        p.drawLine(QPointF(x, y), QPointF(x + s * 0.02, y - s * 0.2))
        p.drawLine(QPointF(x, y), QPointF(x - s * 0.19, y - s * 0.04))
    elif kind == "app":
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#1e63e9"))
        p.drawRoundedRect(QRectF(0, 0, s, s), s * 0.22, s * 0.22)
        pen2 = QPen(QColor("white"))
        pen2.setWidthF(s / 7)
        pen2.setCapStyle(Qt.RoundCap)
        p.setPen(pen2)
        p.drawLine(QPointF(s * 0.32, s * 0.68), QPointF(s * 0.68, s * 0.32))
    p.end()
    return pm


def make_empty_illustration(size=150):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    s = size
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#e9eef8"))
    p.drawEllipse(QRectF(s * 0.05, s * 0.05, s * 0.8, s * 0.8))
    p.setBrush(QColor("white"))
    p.setPen(QPen(QColor("#8a94a8"), 2))
    p.drawRoundedRect(QRectF(s * 0.27, s * 0.2, s * 0.36, s * 0.46), 4, 4)
    p.setPen(QPen(QColor("#a3acbd"), 2.4))
    for i in range(4):
        y = s * (0.3 + i * 0.08)
        p.drawLine(QPointF(s * 0.34, y), QPointF(s * 0.55, y))
    p.setBrush(QColor("#e9eef8"))
    p.setPen(QPen(QColor("#1c2b4a"), 4))
    p.drawEllipse(QRectF(s * 0.5, s * 0.5, s * 0.2, s * 0.2))
    pen = QPen(QColor("#1e63e9"), 5)
    pen.setCapStyle(Qt.RoundCap)
    p.setPen(pen)
    p.drawLine(QPointF(s * 0.68, s * 0.68), QPointF(s * 0.78, s * 0.78))
    p.end()
    return pm


class _Popup(QFrame):
    closed = Signal()

    def __init__(self, parent):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.setObjectName("ddPopup")

    def hideEvent(self, e):
        super().hideEvent(e)
        self.closed.emit()


class MultiSelectDropdown(QFrame):
    selectionChanged = Signal()

    def __init__(self, placeholder, parent=None):
        super().__init__(parent)
        self.setObjectName("ddField")
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(46)
        self.placeholder = placeholder
        self.options = []         
        self.selected = set()     
        self._last_close = 0.0

        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 0, 12, 0)
        lay.setSpacing(10)
        icon = QLabel()
        icon.setPixmap(make_icon("search"))
        self.text_lbl = QLabel()
        self.text_lbl.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        chev = QLabel()
        chev.setPixmap(make_icon("chevron", "#1c2b4a"))
        lay.addWidget(icon)
        lay.addWidget(self.text_lbl, 1)
        lay.addWidget(chev)

        # popup
        self.popup = _Popup(self)
        pl = QVBoxLayout(self.popup)
        pl.setContentsMargins(6, 6, 6, 6)
        pl.setSpacing(6)
        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("ddSearch")
        self.search_edit.setPlaceholderText("Type to filter...")
        self.search_edit.addAction(QIcon(make_icon("search")), QLineEdit.LeadingPosition)
        self.search_edit.textChanged.connect(self._apply_filter)
        self.list = QListWidget()
        self.list.setObjectName("ddList")
        self.list.setSelectionMode(QAbstractItemView.NoSelection)
        self.list.setFocusPolicy(Qt.NoFocus)
        self.list.itemClicked.connect(self._toggle_item)
        pl.addWidget(self.search_edit)
        pl.addWidget(self.list)
        self.popup.closed.connect(self._on_popup_closed)

        self._update_text()

    def set_options(self, values):
        values = sorted(values, key=lambda v: str(v).lower())
        self.options = values
        valid = set(values)
        old = set(self.selected)
        self.selected &= valid
        self._rebuild_list()
        self._update_text()
        return self.selected != old

    def clear_selection(self):
        self.selected.clear()
        self._rebuild_list()
        self._update_text()


    def _rebuild_list(self):
        self.list.clear()
        for v in self.options:
            item = QListWidgetItem(fmt_value(v))
            item.setData(Qt.UserRole, v)
            item.setFlags(Qt.ItemIsEnabled)  
            self._paint_item(item, v in self.selected)
            self.list.addItem(item)
        self._apply_filter(self.search_edit.text())

    def _paint_item(self, item, checked):
        item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        item.setBackground(QBrush(QColor("#eaf1fe")) if checked else QBrush(Qt.transparent))

    def _toggle_item(self, item):
        v = item.data(Qt.UserRole)
        if v in self.selected:
            self.selected.discard(v)
            self._paint_item(item, False)
        else:
            self.selected.add(v)
            self._paint_item(item, True)
        self._update_text()
        self.selectionChanged.emit()

    def _apply_filter(self, text):
        t = text.strip().lower()
        for i in range(self.list.count()):
            it = self.list.item(i)
            it.setHidden(bool(t) and t not in it.text().lower())

    def _update_text(self):
        if not self.selected:
            text = self.placeholder
            self.text_lbl.setProperty("placeholder", True)
        else:
            names = [fmt_value(v) for v in self.options if v in self.selected]
            text = ", ".join(names)
            self.text_lbl.setProperty("placeholder", False)
        self.text_lbl.style().unpolish(self.text_lbl)
        self.text_lbl.style().polish(self.text_lbl)
        fm = self.text_lbl.fontMetrics()
        w = max(40, self.text_lbl.width())
        self.text_lbl.setText(fm.elidedText(text, Qt.ElideRight, w))
        self.text_lbl.setToolTip(text if self.selected else "")

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._update_text()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            if time.monotonic() - self._last_close < 0.25:
                return
            self._open_popup()

    def _open_popup(self):
        self._set_open(True)
        rows = min(max(self.list.count(), 1), 8)
        self.popup.setFixedWidth(self.width())
        self.popup.setFixedHeight(56 + rows * 40)
        self.popup.move(self.mapToGlobal(QPoint(0, self.height() + 4)))
        self.search_edit.clear()
        self.popup.show()
        self.search_edit.setFocus()

    def _on_popup_closed(self):
        self._last_close = time.monotonic()
        self._set_open(False)

    def _set_open(self, is_open):
        self.setProperty("open", is_open)
        self.style().unpolish(self)
        self.style().polish(self)



class LoginScreen(QWidget):
    connected = Signal(object, str, str, str, str)  

    def __init__(self):
        super().__init__()
        self.server_conn = None
        self.loaded_server = None

        outer = QVBoxLayout(self)
        outer.setAlignment(Qt.AlignCenter)

        card = QFrame()
        card.setObjectName("card")
        card.setFixedWidth(540)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(36, 32, 36, 32)
        lay.setSpacing(8)

        title = QLabel("Connect to SQL Server")
        title.setObjectName("title")
        sub = QLabel("Windows Authentication")
        sub.setObjectName("subtitle")
        lay.addWidget(title)
        lay.addWidget(sub)
        lay.addSpacing(16)

        lay.addWidget(self._label("Server Name"))
        row = QHBoxLayout()
        row.setSpacing(10)
        self.server_edit = QLineEdit()
        self.server_edit.setPlaceholderText(r"e.g. localhost\SQLEXPRESS")
        self.server_edit.editingFinished.connect(lambda: self.load_databases(force=False))
        self.load_btn = QPushButton("  Load")
        self.load_btn.setObjectName("secondary")
        self.load_btn.setIcon(QIcon(make_icon("refresh", "#1c2b4a")))
        self.load_btn.setCursor(Qt.PointingHandCursor)
        self.load_btn.clicked.connect(lambda: self.load_databases(force=True))
        row.addWidget(self.server_edit, 1)
        row.addWidget(self.load_btn)
        lay.addLayout(row)
        lay.addSpacing(10)

        lay.addWidget(self._label("Database"))
        self.db_combo = QComboBox()
        self.db_combo.setPlaceholderText("Select database")
        self.db_combo.currentIndexChanged.connect(self.load_tables)
        lay.addWidget(self.db_combo)
        lay.addSpacing(10)

        lay.addWidget(self._label("Table"))
        self.table_combo = QComboBox()
        self.table_combo.setPlaceholderText("Select table")
        lay.addWidget(self.table_combo)

        self.status = QLabel("")
        self.status.setObjectName("status")
        self.status.setWordWrap(True)
        lay.addSpacing(6)
        lay.addWidget(self.status)
        lay.addSpacing(10)

        self.submit_btn = QPushButton("Submit")
        self.submit_btn.setObjectName("primary")
        self.submit_btn.setCursor(Qt.PointingHandCursor)
        self.submit_btn.clicked.connect(self.submit)
        lay.addWidget(self.submit_btn)

        outer.addWidget(card)

    def _label(self, text):
        l = QLabel(text)
        l.setObjectName("fieldLabel")
        return l

    def _error(self, msg):
        QMessageBox.critical(self, "Error", str(msg))

    def load_databases(self, force=False):
        server = self.server_edit.text().strip()
        if not server:
            if force:
                self._error("Enter the server name first.")
            return
        if not force and server == self.loaded_server:
            return

        self.status.setText("Connecting...")
        QApplication.processEvents()
        with busy():
            try:
                if self.server_conn:
                    self.server_conn.close()
                self.server_conn = db_connect(server, "master")
                cur = self.server_conn.cursor()
                cur.execute(
                    "SELECT name FROM sys.databases "
                    "WHERE database_id > 4 AND state_desc = 'ONLINE' "
                    "AND HAS_DBACCESS(name) = 1 ORDER BY name"
                )
                dbs = [r[0] for r in cur.fetchall()]
            except Exception as e:
                self.server_conn = None
                self.loaded_server = None
                self.status.setText("")
                self._error(f"Could not connect to '{server}'.\n\n{e}")
                return

        self.loaded_server = server
        self.db_combo.blockSignals(True)
        self.db_combo.clear()
        self.db_combo.addItems(dbs)
        self.db_combo.setCurrentIndex(-1)
        self.db_combo.blockSignals(False)
        self.table_combo.clear()
        self.status.setText(f"Connected. {len(dbs)} database(s) found.")

    def load_tables(self, idx):
        self.table_combo.clear()
        if idx < 0 or not self.server_conn:
            return
        db = self.db_combo.currentText()
        with busy():
            try:
                cur = self.server_conn.cursor()
                cur.execute(
                    f"SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE "
                    f"FROM {q(db)}.INFORMATION_SCHEMA.TABLES "
                    f"ORDER BY TABLE_SCHEMA, TABLE_NAME"
                )
                rows = cur.fetchall()
            except Exception as e:
                self._error(f"Could not read tables from '{db}'.\n\n{e}")
                return
        for schema, name, ttype in rows:
            label = f"{schema}.{name}" + ("  (view)" if ttype == "VIEW" else "")
            self.table_combo.addItem(label, (schema, name))
        self.table_combo.setCurrentIndex(-1)
        self.status.setText(f"{len(rows)} table(s) in {db}.")

    def submit(self):
        server = self.server_edit.text().strip()
        db = self.db_combo.currentText()
        tdata = self.table_combo.currentData()
        if not server or not db or not tdata:
            self._error("Please enter the server and select a database and table.")
            return
        schema, table = tdata
        full = f"{q(schema)}.{q(table)}"

        with busy():
            try:
                conn = db_connect(server, db)
                cur = conn.cursor()
                cur.execute(f"SELECT TOP 0 * FROM {full}")
                cols = [d[0] for d in cur.description]
            except Exception as e:
                self._error(f"Could not open table.\n\n{e}")
                return

        lower = {c.lower() for c in cols}
        needed = [USER_COLUMN, CLASS_COLUMN, FLAG_COLUMN] + [c for c, _ in RESULT_COLUMNS]
        missing = [c for c in dict.fromkeys(needed) if c.lower() not in lower]
        if missing:
            conn.close()
            self._error(
                "These columns from the CONFIG section were not found in the table:\n  "
                + ", ".join(missing)
                + "\n\nColumns available in this table:\n  " + ", ".join(cols)
                + "\n\nUpdate the column names at the top of the script."
            )
            return

        self.connected.emit(conn, server, db, schema, table)

    def close_conn(self):
        if self.server_conn:
            try:
                self.server_conn.close()
            except Exception:
                pass
            self.server_conn = None


class FilterScreen(QWidget):
    back_requested = Signal()

    def __init__(self, conn, schema, table, info_text=""):
        super().__init__()
        self.conn = conn
        self.table_sql = f"{q(schema)}.{q(table)}"
        self.rows = []  # distinct (user, class, flag) combos

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---------- top filter area ----------
        top = QWidget()
        top.setObjectName("topArea")
        tl = QVBoxLayout(top)
        tl.setContentsMargins(40, 28, 40, 24)
        tl.setSpacing(18)

        grid = QGridLayout()
        grid.setHorizontalSpacing(24)
        grid.setVerticalSpacing(10)
        self.dd_user = MultiSelectDropdown(USER_PLACEHOLDER)
        self.dd_class = MultiSelectDropdown(CLASS_PLACEHOLDER)
        self.dd_flag = MultiSelectDropdown(FLAG_PLACEHOLDER)
        self.dropdowns = [self.dd_user, self.dd_class, self.dd_flag]
        for col, (lbl, dd) in enumerate(zip([USER_LABEL, CLASS_LABEL, FLAG_LABEL], self.dropdowns)):
            l = QLabel(lbl)
            l.setObjectName("fieldLabel")
            grid.addWidget(l, 0, col)
            grid.addWidget(dd, 1, col)
            grid.setColumnStretch(col, 1)
            dd.selectionChanged.connect(self.refresh_options)
        tl.addLayout(grid)

        btn_row = QHBoxLayout()
        back = QPushButton("← Change connection")
        back.setObjectName("link")
        back.setCursor(Qt.PointingHandCursor)
        back.clicked.connect(self.back_requested.emit)
        info = QLabel(info_text)
        info.setObjectName("subtitle")
        self.search_btn = QPushButton("  Search")
        self.search_btn.setObjectName("primary")
        self.search_btn.setIcon(QIcon(make_icon("search", "white")))
        self.search_btn.setCursor(Qt.PointingHandCursor)
        self.search_btn.setMinimumWidth(170)
        self.search_btn.clicked.connect(self.run_search)
        self.clear_btn = QPushButton("  Clear")
        self.clear_btn.setObjectName("secondary")
        self.clear_btn.setIcon(QIcon(make_icon("refresh", "#1c2b4a")))
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setMinimumWidth(160)
        self.clear_btn.clicked.connect(self.clear_all)
        btn_row.addWidget(back)
        btn_row.addSpacing(12)
        btn_row.addWidget(info)
        btn_row.addStretch(1)
        btn_row.addWidget(self.search_btn)
        btn_row.addSpacing(8)
        btn_row.addWidget(self.clear_btn)
        tl.addLayout(btn_row)
        root.addWidget(top)

        line = QFrame()
        line.setObjectName("divider")
        line.setFixedHeight(1)
        root.addWidget(line)

        # ---------- results card ----------
        body = QWidget()
        bl = QVBoxLayout(body)
        bl.setContentsMargins(40, 32, 40, 32)
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)

        head = QHBoxLayout()
        head.setContentsMargins(26, 16, 26, 14)
        rt = QLabel("Results")
        rt.setObjectName("resultsTitle")
        self.count_lbl = QLabel("")
        self.count_lbl.setObjectName("subtitle")
        head.addWidget(rt)
        head.addStretch(1)
        head.addWidget(self.count_lbl)
        cl.addLayout(head)

        self.table = QTableWidget(0, len(RESULT_COLUMNS))
        self.table.setHorizontalHeaderLabels([h for _, h in RESULT_COLUMNS])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setSortingEnabled(True)

        # empty state
        self.empty = QWidget()
        el = QVBoxLayout(self.empty)
        el.setAlignment(Qt.AlignCenter)
        img = QLabel()
        img.setPixmap(make_empty_illustration())
        img.setAlignment(Qt.AlignCenter)
        self.empty_title = QLabel("No results to display")
        self.empty_title.setObjectName("emptyTitle")
        self.empty_title.setAlignment(Qt.AlignCenter)
        self.empty_sub = QLabel("Adjust your search criteria and try again.")
        self.empty_sub.setObjectName("emptySub")
        self.empty_sub.setAlignment(Qt.AlignCenter)
        el.addWidget(img)
        el.addWidget(self.empty_title)
        el.addWidget(self.empty_sub)

        # header row stays visible even when empty (like the design)
        self.header_only = QTableWidget(0, len(RESULT_COLUMNS))
        self.header_only.setHorizontalHeaderLabels([h for _, h in RESULT_COLUMNS])
        self.header_only.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.header_only.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.header_only.verticalHeader().setVisible(False)
        self.header_only.setFixedHeight(self.header_only.horizontalHeader().sizeHint().height())
        self.header_only.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.header_only.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.header_only.setEnabled(False)

        empty_wrap = QWidget()
        empty_wrap.setObjectName("emptyWrap")
        empty_wrap.setAttribute(Qt.WA_StyledBackground, True)
        ew = QVBoxLayout(empty_wrap)
        ew.setContentsMargins(0, 0, 0, 0)
        ew.setSpacing(0)
        ew.addWidget(self.header_only)
        ew.addWidget(self.empty, 1)

        self.stack = QStackedWidget()
        self.stack.setObjectName("resultsStack")
        self.stack.addWidget(empty_wrap)   # 0
        self.stack.addWidget(self.table)   # 1
        cl.addWidget(self.stack, 1)

        bl.addWidget(card)
        root.addWidget(body, 1)

        self.load_data()

    # ---- data ----
    def load_data(self):
        sql = (f"SELECT DISTINCT {q(USER_COLUMN)}, {q(CLASS_COLUMN)}, {q(FLAG_COLUMN)} "
               f"FROM {self.table_sql}")
        with busy():
            try:
                cur = self.conn.cursor()
                cur.execute(sql)
                self.rows = [tuple(r) for r in cur.fetchall()]
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load dropdown values.\n\n{e}")
                self.rows = []
        self.refresh_options()

    def refresh_options(self):
        """Each dropdown only shows values that match what's selected in the other two."""
        for _ in range(4):  # repeat until selections are stable
            sel = [set(d.selected) for d in self.dropdowns]

            def match(row, skip):
                return all(i == skip or not sel[i] or row[i] in sel[i] for i in range(3))

            changed = False
            for i, dd in enumerate(self.dropdowns):
                values = {r[i] for r in self.rows if r[i] is not None and match(r, i)}
                if dd.set_options(values):
                    changed = True
            if not changed:
                break

    def run_search(self):
        cols = ", ".join(q(c) for c, _ in RESULT_COLUMNS)
        where, params = [], []
        for col, dd in zip([USER_COLUMN, CLASS_COLUMN, FLAG_COLUMN], self.dropdowns):
            if dd.selected:
                vals = list(dd.selected)
                where.append(f"{q(col)} IN ({', '.join('?' * len(vals))})")
                params.extend(vals)
        sql = f"SELECT TOP ({MAX_RESULT_ROWS}) {cols} FROM {self.table_sql}"
        if where:
            sql += " WHERE " + " AND ".join(where)

        with busy():
            try:
                cur = self.conn.cursor()
                cur.execute(sql, params)
                data = cur.fetchall()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Search failed.\n\n{e}")
                return
        self.show_results(data)

    def show_results(self, data):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        if not data:
            self.count_lbl.setText("0 rows")
            self.stack.setCurrentIndex(0)
            return
        self.table.setRowCount(len(data))
        for r, row in enumerate(data):
            for c, v in enumerate(row):
                item = QTableWidgetItem(fmt_value(v))
                if isinstance(v, (int, float, decimal.Decimal)) and not isinstance(v, bool):
                    item.setData(Qt.EditRole, float(v))
                    item.setText(fmt_value(v))
                self.table.setItem(r, c, item)
        self.table.setSortingEnabled(True)
        n = len(data)
        self.count_lbl.setText(f"{n} row{'s' if n != 1 else ''}"
                               + (f" (limit {MAX_RESULT_ROWS})" if n >= MAX_RESULT_ROWS else ""))
        self.stack.setCurrentIndex(1)

    def clear_all(self):
        for dd in self.dropdowns:
            dd.clear_selection()
        self.refresh_options()
        self.table.setRowCount(0)
        self.count_lbl.setText("")
        self.stack.setCurrentIndex(0)

    def close_conn(self):
        try:
            self.conn.close()
        except Exception:
            pass


# ---------------------------------------------------------------------
#  Main window
# ---------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(QIcon(make_icon("app", size=32)))
        self.resize(1240, 820)

        self.stack = QStackedWidget()
        self.stack.setObjectName("root")
        self.setCentralWidget(self.stack)

        self.login = LoginScreen()
        self.login.connected.connect(self.open_filter_screen)
        self.stack.addWidget(self.login)
        self.filter_screen = None

    def open_filter_screen(self, conn, server, db, schema, table):
        self._drop_filter_screen()
        self.filter_screen = FilterScreen(conn, schema, table, f"{server} · {db} · {schema}.{table}")
        self.filter_screen.back_requested.connect(self.go_back)
        self.stack.addWidget(self.filter_screen)
        self.stack.setCurrentWidget(self.filter_screen)

    def go_back(self):
        self._drop_filter_screen()
        self.stack.setCurrentWidget(self.login)

    def _drop_filter_screen(self):
        if self.filter_screen:
            self.filter_screen.close_conn()
            self.stack.removeWidget(self.filter_screen)
            self.filter_screen.deleteLater()
            self.filter_screen = None

    def closeEvent(self, e):
        self._drop_filter_screen()
        self.login.close_conn()
        super().closeEvent(e)


STYLE = """
* { font-family: 'Segoe UI', 'Inter', Arial; font-size: 14px; color: #1c2b4a; }
QMainWindow, #root { background: #f4f7fc; }
#resultsStack, #emptyWrap { background: white; border-bottom-left-radius: 10px; border-bottom-right-radius: 10px; }
#topArea { background: #f8faff; }
#divider { background: #e1e7f2; }

#card { background: white; border: 1px solid #dfe5f0; border-radius: 10px; }
#title { font-size: 22px; font-weight: 700; }
#subtitle { color: #6b7590; font-size: 13px; }
#status { color: #4a6fb5; font-size: 13px; }
#fieldLabel { font-size: 15px; font-weight: 600; }
#resultsTitle { font-size: 24px; font-weight: 700; }
#emptyTitle { font-size: 20px; font-weight: 700; margin-top: 6px; }
#emptySub { color: #5b6478; font-size: 15px; }

QLineEdit, QComboBox {
    background: white; border: 1px solid #cfd7e6; border-radius: 6px;
    padding: 0 12px; min-height: 42px;
}
QLineEdit:focus, QComboBox:focus { border: 1px solid #1e63e9; }
QComboBox::drop-down { border: none; width: 30px; }
QComboBox QAbstractItemView {
    background: white; border: 1px solid #cfd7e6; selection-background-color: #eaf1fe;
    selection-color: #1c2b4a; outline: 0; padding: 4px;
}

#ddField { background: white; border: 1px solid #cfd7e6; border-radius: 6px; }
#ddField:hover { border: 1px solid #9fb3d9; }
#ddField[open="true"] { border: 1px solid #1e63e9; }
QLabel[placeholder="true"] { color: #4b5468; }

#ddPopup { background: white; border: 1px solid #cfd7e6; border-radius: 6px; }
#ddSearch { min-height: 34px; border: 1px solid #e1e7f2; }
#ddList { border: none; outline: 0; background: white; }
#ddList::item { padding: 8px 6px; border-radius: 4px; margin: 1px 0; }
#ddList::item:hover { background: #f2f6fd; }

QPushButton { border-radius: 6px; padding: 0 22px; min-height: 46px; font-size: 16px; font-weight: 600; }
QPushButton#primary { background: #1e63e9; color: white; border: none; }
QPushButton#primary:hover { background: #1853c7; }
QPushButton#primary:pressed { background: #1446a8; }
QPushButton#secondary { background: white; border: 1px solid #cfd7e6; }
QPushButton#secondary:hover { background: #f2f6fd; }
QPushButton#link { background: transparent; border: none; color: #1e63e9; font-size: 13px;
                   padding: 0; min-height: 0; font-weight: 600; }
QPushButton#link:hover { text-decoration: underline; }

QTableWidget { border: none; background: white; alternate-background-color: #fafbfe;
               gridline-color: #eef1f7; selection-background-color: #eaf1fe; selection-color: #1c2b4a; }
QTableWidget::item { padding: 0 14px; border-bottom: 1px solid #eef1f7; }
QHeaderView::section {
    background: #eef2fa; color: #1c2b4a; font-size: 15px; font-weight: 700;
    padding: 14px 18px; border: none; border-right: 1px solid #dfe5f0;
    border-top: 1px solid #dfe5f0; border-bottom: 1px solid #dfe5f0;
}
QHeaderView::section:last { border-right: none; }
QScrollBar:vertical { background: transparent; width: 10px; }
QScrollBar::handle:vertical { background: #cfd7e6; border-radius: 5px; min-height: 30px; }
QScrollBar::add-line, QScrollBar::sub-line { height: 0; }
"""


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
