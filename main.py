# main.py (PySide6 version with tabs + responsive flow, no blinking)

import os, sys, json, time, traceback, importlib, re, threading, queue
from pathlib import Path

# ---- PySide6 / Qt ----
from PySide6.QtCore import Qt, QTimer, QPointF, QSize
from PySide6.QtGui import (
    QIcon, QColor, QBrush, QPen, QPolygonF, QPixmap, QPainterPath, QPainter
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QCheckBox, QScrollArea, QFrame,
    QProgressBar, QPlainTextEdit, QGridLayout,
    QGraphicsView, QGraphicsScene, QGraphicsPolygonItem, QGraphicsPixmapItem,
    QGraphicsPathItem, QGraphicsDropShadowEffect
)

# COM (Outlook/Excel) init for worker threads
import pythoncom

# ------------------ SIMPLE SETTINGS PERSISTENCE ------------------
SETTINGS_FILE = Path(__file__).with_name("user_settings.json")

def _load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_settings(data: dict) -> None:
    try:
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass

def _apply_treat_env(username: str, password: str) -> None:
    os.environ["TREAT_USERNAME"] = username
    os.environ["TREAT_PASSWORD"] = password

# ------------------ PIPELINE STEPS ------------------
STEPS = [
    ("clean",  "[1/6] Cleaning Folder",               "Clean_Download_Folder"),
    ("treat",  "[2/6] Pulling Treat Reports",         "Treat_Pulling"),
    ("move",   "[3/6] Move Treat files to One-Drive", "Treat_File_Moving"),
    ("alaya",  "[5/6] Pulling AlayaCare Reports",     "AlayaCare_Pulling"),
    ("tclean", "[4/6] Cleaning Treat Reports",        "Treat_Data_Cleaning"),
    ("aclean", "[6/6] AlayaCare Data Cleaning",       "AlayaCare_Data_Cleaning"),
]

# ------------------ REPORTS ------------------
REPORTS = [
    ("hhri", "HHRI (Save to Downloads)", "HHRI_Hours"),
    ("fame", "CMHA Peel – FAME Monthly (save to Downloads)", "FAME_Report"),
    ("jam",  "JAM Report (Save to Downloads)", "JAM_Report"),
    ("ocan","Overdue OCAN List (Save to Downloads)", "Overdue_OCAN_List")
]

STATUS_COLORS = {
    "pending": "#666666",
    "running": "#1f6feb",
    "done":    "#18794e",
    "fail":    "#b54708",
}

# ---------- Visual labels for the flow nodes (not the row titles) ----------
FLOW_LABELS = {
    "clean":        "Clean Folder",
    "treat_box":    "Treat",
    "raw":          "Raw Data",
    "alayacare":    "AlayaCare",
    "clean_data":   "Clean Data",
    # keeping these in case you still want them visible
    "downloads":    "Downloads",
    "email":        "Email Inbox",
}

# Map pipeline step -> which node(s) should light up
STEP_NODE_MAP = {
    "clean":  ["clean"],
    "treat":  ["treat_box","downloads"],
    "move":   ["raw"],
    "alaya":  ["alayacare","email","raw"],
    "tclean": [],
    "aclean": ["clean_data"],
}

# Which edges should change per pipeline step
STEP_EDGE_MAP = {
    "clean":  [("clean","treat_box")],
    "treat":  [("treat_box", "downloads")],
    "move":   [("downloads", "raw")],
    "alaya":  [("alayacare", "email"),("email","raw")],
    "tclean": [],  # stays green when cleaning
    "aclean": [("raw", "clean_data")],  # same node reused
}

# All edges in the diagram (used for resetting)
ALL_EDGES = [
    ("clean", "treat_box"),
    ("treat_box", "downloads"),
    ("downloads","raw"),
    ("alayacare","email"),
    ("email", "raw"),
    ("raw", "clean_data"),
]


# ------------------ Helpers ------------------
def _step_parts(title_text: str):
    m = re.match(r"^\[(\d+/\d+)\]\s*(.+)$", title_text.strip())
    return (m.group(1), m.group(2)) if m else (None, title_text)

def _get_runner(module_name: str):
    # reload so modules re-read env vars each step
    if module_name in sys.modules:
        mod = importlib.reload(sys.modules[module_name])
    else:
        mod = importlib.import_module(module_name)
    if hasattr(mod, "run") and callable(mod.run):  return mod.run
    if hasattr(mod, "main") and callable(mod.main): return mod.main
    raise AttributeError(f"{module_name} has no callable run() or main().")

# ------------------ Qt FlowView with optional picture nodes ------------------
class FlowView(QGraphicsView):
    # --- Responsive layout baseline (design-space) ---
    DESIGN_W = 1800    # your original coordinate width
    DESIGN_H = 360     # your original coordinate height

    DEFAULT_ICON_SIZE = 220
    BASE_ICON_SIZES = {
        "clean": 220, "treat_box": 400, "alayacare": 400,
        "downloads": 180, "email": 180, "raw": 220, "clean_data": 220,
    }
    BASE_NODE_POS = {
        "clean": (180, 60),
        "treat_box": (600, 60),
        "alayacare": (600, 270),
        "downloads": (1000, 60),
        "email":     (1000, 270),
        "raw": (1350, 160),
        "clean_data": (1750, 160),
    }
    SCENE_PADDING = 60  # extra breathing room around everything

    def __init__(self, parent=None, icons_dir: Path | None = None):
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self.setScene(self._scene)
        # Correct render hints: use QPainter flags, not QPainterPath
        self.setRenderHint(QPainter.Antialiasing, True)
        self.setRenderHint(QPainter.TextAntialiasing, True)
        self.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.setFrameShape(QFrame.NoFrame)
        self.setMinimumHeight(300)

        self.icons_dir = icons_dir or (Path(__file__).resolve().parent / "icons")

        # state storage
        self._nodes = {}   # key -> {"pix": item, "bbox": (x0,y0,x1,y1), "effect": effect}
        self._edges = []   # list of dicts per edge (see _add_arrow)
        self._states = {
            "clean": "pending", "treat_box": "pending", "alayacare": "pending",
            "downloads": "pending", "email": "pending", "raw": "pending", "clean_data": "pending",
        }

        self.rebuild()

    # ---------- responsive helpers ----------
    def _scale_factor(self) -> float:
        # Fit the design-space into the current viewport, keep aspect ratio
        vw = max(1, self.viewport().width()  - 2*self.SCENE_PADDING)
        vh = max(1, self.viewport().height() - 2*self.SCENE_PADDING)
        sx = vw / self.DESIGN_W
        sy = vh / self.DESIGN_H
        s = min(sx, sy)
        # clamp to keep icons readable (tweak as you like)
        return max(0.35, min(s, 2.0))

    def _scaled_icon_size(self, key: str, s: float) -> int:
        base = self.BASE_ICON_SIZES.get(key, self.DEFAULT_ICON_SIZE)
        return int(max(24, base * s))

    def _scaled_pos(self, key: str, s: float) -> tuple[float, float]:
        x, y = self.BASE_NODE_POS[key]
        return (x * s, y * s)

    # ---------- utilities ----------
    def _icon_for(self, key: str) -> QPixmap | None:
        p = self.icons_dir / f"{key}.png"
        if p.exists():
            pm = QPixmap(str(p))
            if not pm.isNull():
                return pm
        return None

    def _scaled(self, pm: QPixmap, size_px: int) -> QPixmap:
        size_px = int(max(24, min(2048, size_px)))
        return pm.scaled(QSize(size_px, size_px), Qt.KeepAspectRatio, Qt.SmoothTransformation)

    # ---------- node/edge drawing ----------
    def _add_node_image(self, key: str, cx: float, cy: float, size_px: int):
        pm = self._icon_for(key)
        if pm is None:
            pm = QPixmap(size_px, size_px); pm.fill(QColor("#e8e8e8"))
        spm = self._scaled(pm, size_px)
        pix = QGraphicsPixmapItem(spm)
        pix.setZValue(10)
        self._scene.addItem(pix)

        # center at (cx, cy)
        x = cx - spm.width() / 2
        y = cy - spm.height() / 2
        pix.setPos(x, y)
        bbox = (x, y, x + spm.width(), y + spm.height())

        eff = QGraphicsDropShadowEffect()
        eff.setBlurRadius(0); eff.setOffset(0, 0); eff.setColor(QColor(0, 0, 0, 0))
        pix.setGraphicsEffect(eff)

        self._nodes[key] = {"pix": pix, "bbox": bbox, "effect": eff}

    def _arrow_points(self, from_key, to_key):
        fx0, fy0, fx1, fy1 = self._nodes[from_key]["bbox"]
        tx0, ty0, tx1, ty1 = self._nodes[to_key]["bbox"]
        start = QPointF(fx1, (fy0 + fy1) / 2.0)
        end   = QPointF(tx0, (ty0 + ty1) / 2.0)
        return start, end

    # ---- EDGE state helpers ----
    def _edge_state_color(self, state: str) -> QColor:
        if state in ("running", "done"): return QColor("#18794e")
        if state == "fail":              return QColor("#b54708")
        return QColor("#888")

    def _apply_edge_pen(self, e, color: QColor, width: int = 4):
        pen = QPen(color, width)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        e["path"].setPen(pen)
        e["head"].setBrush(QBrush(color))
        e["head"].setPen(QPen(color))

    def set_edge_state(self, from_key: str, to_key: str, state: str):
        for e in self._edges:
            if e["from"] == from_key and e["to"] == to_key:
                e["state"] = state
                col = self._edge_state_color(state)
                self._apply_edge_pen(e, col, width=4)

                # steady glow (no animation)
                for eff_key in ("path_effect", "head_effect"):
                    eff = e.get(eff_key)
                    if eff is None:
                        eff = QGraphicsDropShadowEffect()
                        eff.setOffset(0, 0)
                        if eff_key == "path_effect":
                            e["path"].setGraphicsEffect(eff)
                        else:
                            e["head"].setGraphicsEffect(eff)
                        e[eff_key] = eff

                    if state == "pending":
                        eff.setBlurRadius(0); eff.setColor(QColor(0, 0, 0, 0))
                    else:
                        eff.setBlurRadius(20); eff.setColor(col)
                break

    def _add_arrow(self, from_key, to_key):
        start, end = self._arrow_points(from_key, to_key)

        # Cubic curve for a nicer arrow
        dx = (end.x() - start.x())
        c1 = QPointF(start.x() + dx * 0.35, start.y())
        c2 = QPointF(end.x()   - dx * 0.35, end.y())

        path = QPainterPath(start)
        path.cubicTo(c1, c2, end)

        path_item = QGraphicsPathItem(path)
        base_col = QColor("#888")
        pen = QPen(base_col, 4)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        path_item.setPen(pen)
        path_item.setZValue(5)
        self._scene.addItem(path_item)

        # arrowhead aligned with tangent
        tangent = path.angleAtPercent(1.0)  # degrees
        head_len = 16
        head_wid = 8
        base = end
        from math import radians, cos, sin
        t = radians(-tangent)
        ux, uy = cos(t), sin(t)
        left  = QPointF(base.x() - head_len*ux + head_wid*uy, base.y() - head_len*uy - head_wid*ux)
        right = QPointF(base.x() - head_len*ux - head_wid*uy, base.y() - head_len*uy + head_wid*ux)
        head = QGraphicsPolygonItem(QPolygonF([base, left, right]))
        head.setBrush(QBrush(base_col))
        head.setPen(QPen(base_col))
        head.setZValue(6)
        self._scene.addItem(head)

        self._edges.append({
            "from": from_key,
            "to": to_key,
            "path": path_item,
            "head": head,
            "state": "pending",
            "path_effect": None,
            "head_effect": None,
        })

    # ---------- build / rebuild ----------
    def rebuild(self):
        self._scene.clear()
        self._nodes.clear()
        self._edges.clear()

        s = self._scale_factor()

        # place nodes using scaled positions & sizes
        for key in self.BASE_NODE_POS.keys():
            cx, cy   = self._scaled_pos(key, s)
            size_px  = self._scaled_icon_size(key, s)
            self._add_node_image(key, cx, cy, size_px)

        # arrows (same logical flow)
        self._add_arrow("clean", "treat_box")
        self._add_arrow("treat_box", "downloads")
        self._add_arrow("alayacare", "email")
        self._add_arrow("downloads", "raw")
        self._add_arrow("email", "raw")
        self._add_arrow("raw", "clean_data")

        # re-apply saved states (steady glow)
        for k, st in self._states.items():
            self._apply_state(k, st)

        # Scene rect sized to scaled design, with padding
        pad = self.SCENE_PADDING
        scene_w = self.DESIGN_W * s
        scene_h = self.DESIGN_H * s
        self._scene.setSceneRect(-pad, -pad, scene_w + 2*pad, scene_h + 2*pad)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        # Rebuild after resize; singleShot avoids thrashing while dragging
        QTimer.singleShot(0, self.rebuild)

    # ---------- node state (steady glow) ----------
    def _state_color(self, state: str) -> QColor:
        if state in ("running", "done"): return QColor("#18794e")   # GREEN
        if state == "fail":              return QColor("#b54708")
        return QColor(0, 0, 0, 0)

    def _apply_state(self, key, state):
        self._states[key] = state
        node = self._nodes.get(key)
        if not node: return
        eff: QGraphicsDropShadowEffect = node["effect"]
        col = self._state_color(state)
        if state == "pending":
            eff.setBlurRadius(0); eff.setColor(QColor(0, 0, 0, 0))
        else:
            eff.setBlurRadius(28); eff.setColor(col)

    # external API used by MainWindow
    def set_node_state(self, key, state):
        self._apply_state(key, state)

# ------------------ Main Window ------------------
class MainWindow(QMainWindow):
    
    def _toggle_creds(self):
        self.creds_group.setVisible(not self.creds_group.isVisible())

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Daily Automation")
        self.resize(1200, 800)
        icon_path = Path(__file__).with_name("app.ico")
        if icon_path.exists(): self.setWindowIcon(QIcon(str(icon_path)))

        # settings/env
        self.settings = _load_settings()
        if "TREAT_USERNAME" in self.settings:
            os.environ["TREAT_USERNAME"] = self.settings["TREAT_USERNAME"]
        if "TREAT_PASSWORD" in self.settings:
            os.environ["TREAT_PASSWORD"] = self.settings["TREAT_PASSWORD"]

        # thread queue
        self.q = queue.Queue()
        self.running = False

        # --- central layout ---
        central = QWidget(self)
        self.setCentralWidget(central)
        v = QVBoxLayout(central)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(8)

        # Tabs (Data Automation, Reports)
        self.tabs = QTabWidget()
        v.addWidget(self.tabs, 1)

        self.tab_auto = QWidget()
        self.tab_rep  = QWidget()
        self.tabs.addTab(self.tab_auto, "Data Automation")
        self.tabs.addTab(self.tab_rep, "Reports")

        # ---- Data Automation tab ----
        auto_layout = QVBoxLayout(self.tab_auto)
        auto_layout.setSpacing(8)

        # Header with Run
        hdr = QHBoxLayout()
        back_lbl = QLabel("Data Automation")
        back_lbl.setStyleSheet("font-size:18px; font-weight:600;")
        hdr.addWidget(back_lbl)
        hdr.addStretch(1)

        self.btn_run = QPushButton("Run")
        self.btn_run.clicked.connect(self.start_pipeline)
        hdr.addWidget(self.btn_run)

        # 🔽 Add your toggle button here
        self.btn_toggle_creds = QPushButton("Update Credentials")
        self.btn_toggle_creds.clicked.connect(self._toggle_creds)
        hdr.addWidget(self.btn_toggle_creds)

        auto_layout.addLayout(hdr)

        # Treat creds (keep a handle so we can toggle visibility)
        self.creds_group = QGroupBox("Treat credentials")
        gl = QGridLayout(self.creds_group); gl.setColumnStretch(1, 1); gl.setColumnStretch(3, 1)
        self.in_user = QLineEdit(os.getenv("TREAT_USERNAME", self.settings.get("TREAT_USERNAME", "yxu")))
        self.in_pass = QLineEdit(os.getenv("TREAT_PASSWORD", self.settings.get("TREAT_PASSWORD", "")))
        self.in_pass.setEchoMode(QLineEdit.Password)
        gl.addWidget(QLabel("Username"), 0, 0); gl.addWidget(self.in_user, 0, 1)
        gl.addWidget(QLabel("Password"), 0, 2); gl.addWidget(self.in_pass, 0, 3)
        self.btn_save_creds = QPushButton("Save")
        self.btn_save_creds.clicked.connect(self._save_creds)
        gl.addWidget(self.btn_save_creds, 0, 4)

        # Start hidden; the toggle button will reveal this
        self.creds_group.setVisible(False)

        auto_layout.addWidget(self.creds_group)


        # Flow (with optional icons)
        self.flow = FlowView(self, icons_dir=Path(__file__).resolve().parent / "icons")
        auto_layout.addWidget(self.flow)

        # Scrollable steps list
        steps_group = QGroupBox("Pipeline steps")
        steps_v = QVBoxLayout(steps_group)
        self.steps_scroll = QScrollArea(); self.steps_scroll.setWidgetResizable(True)
        steps_widget = QWidget(); self.steps_scroll.setWidget(steps_widget)
        self.steps_layout = QVBoxLayout(steps_widget)
        self.steps_layout.setContentsMargins(8, 8, 8, 8); self.steps_layout.setSpacing(4)
        self.step_rows = {}
        for key, title_text, _ in STEPS:
            row = QWidget(); hl = QHBoxLayout(row); hl.setContentsMargins(0,0,0,0)
            lbl = QLabel(title_text); lbl.setMinimumWidth(600)
            st  = QLabel("Pending"); st.setStyleSheet(f"color:{STATUS_COLORS['pending']};")
            hl.addWidget(lbl, 1); hl.addWidget(st, 0)
            self.steps_layout.addWidget(row)
            self.step_rows[key] = {"row": row, "label": lbl, "status": st}
        steps_v.addWidget(self.steps_scroll)
        auto_layout.addWidget(steps_group)

        # Controls
        ctl = QHBoxLayout()
        self.progress = QProgressBar(); self.progress.setRange(0,0); self.progress.setVisible(False)
        ctl.addWidget(self.progress); ctl.addStretch(1)
        auto_layout.addLayout(ctl)

        # ---- Reports tab ----
        rep_layout = QVBoxLayout(self.tab_rep)
        rep_layout.setSpacing(8)
        rep_layout.addWidget(QLabel("<b>Reports</b>"))

        self.report_checks = {}
        for key, label, _ in REPORTS:
            cb = QCheckBox(label)
            rep_layout.addWidget(cb)
            self.report_checks[key] = cb

        self.btn_run_reports = QPushButton("Create report(s)")
        self.btn_run_reports.clicked.connect(self.start_reports)
        rep_layout.addWidget(self.btn_run_reports)
        rep_layout.addStretch(1)

        # ---- Shared log at the bottom ----
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(5000)
        v.addWidget(self.log, 0)

        # poll the queue
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._drain_queue)
        self.timer.start(100)

    # ---- creds
    def _save_creds(self):
        u = self.in_user.text().strip()
        p = self.in_pass.text()
        _apply_treat_env(u, p)
        s = dict(_load_settings())
        s["TREAT_USERNAME"] = u
        s["TREAT_PASSWORD"] = p
        _save_settings(s)
        self._log("Saved Treat credentials (env updated).")

    # ---- logging / queue
    def _log(self, msg: str):
        self.log.appendPlainText(msg.rstrip())
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def _drain_queue(self):
        try:
            while True:
                typ, payload = self.q.get_nowait()
                if typ == "status":
                    self._set_status(payload["key"], payload["state"], payload.get("text"))
                elif typ == "log":
                    self._log(payload["msg"])
                elif typ == "done_all":
                    self._set_progress(False)
                    self._log("All steps completed.")
                    self.running = False
                    self.btn_run.setEnabled(True); self.btn_run.setText("Run")
                elif typ == "failed_all":
                    self._set_progress(False)
                    self._log("Pipeline stopped due to failure.")
                    self.running = False
                    self.btn_run.setEnabled(True); self.btn_run.setText("Run")
                elif typ == "reports_done":
                    self._log("Report job(s) completed.")
                    self.btn_run_reports.setEnabled(True); self.btn_run_reports.setText("Create report(s)")
                elif typ == "reports_failed":
                    self._log("Report job(s) stopped due to failure.")
                    self.btn_run_reports.setEnabled(True); self.btn_run_reports.setText("Create report(s)")
        except queue.Empty:
            pass

    # ---- UI helpers
    def _set_progress(self, on: bool):
        self.progress.setVisible(on)

    def _auto_scroll_to_step(self, step_key: str):
        row = self.step_rows.get(step_key, {}).get("row")
        if row:
            self.steps_scroll.ensureWidgetVisible(row, xMargin=0, yMargin=24)

    def _set_status(self, step_key: str, state: str, text: str | None = None):
        # row text/color
        r = self.step_rows.get(step_key)
        if r:
            r["status"].setText(text if text is not None else state.capitalize())
            r["status"].setStyleSheet(f"color:{STATUS_COLORS.get(state, '#666')};")

        if state == "running":
            self._set_progress(True)
            self._auto_scroll_to_step(step_key)

        # ----- NODES (steady glow) -----
        for node_key in STEP_NODE_MAP.get(step_key, []):
            if step_key == "alaya" and state == "running":
                self.flow.set_node_state(node_key, "done")
            else:
                self.flow.set_node_state(node_key, state)

        # ----- EDGES (steady glow) -----
        edges = STEP_EDGE_MAP.get(step_key, [])
        for a, b in edges:
            self.flow.set_edge_state(a, b, state)

    # ---- Actions
    def start_pipeline(self):
        if self.running: return
        # reset UI rows
        for k, r in self.step_rows.items():
            r["status"].setText("Pending")
            r["status"].setStyleSheet(f"color:{STATUS_COLORS['pending']};")

        # reset nodes
        for node_key in FLOW_LABELS:
            self.flow.set_node_state(node_key, "pending")

        # reset edges
        for a, b in ALL_EDGES:
            self.flow.set_edge_state(a, b, "pending")

        self.running = True
        self.btn_run.setEnabled(False); self.btn_run.setText("Running…")
        threading.Thread(target=run_pipeline, args=(self,), daemon=True).start()

    def start_reports(self):
        selected = [(key, label, module) for (key, label, module) in REPORTS if self.report_checks[key].isChecked()]
        if not selected:
            self._log("No reports selected."); return
        self.btn_run_reports.setEnabled(False); self.btn_run_reports.setText("Running…")
        threading.Thread(target=run_reports, args=(self, selected), daemon=True).start()

# ------------------ Workers (same logic as before) ------------------
def run_pipeline(win: MainWindow, stop_on_error: bool = True):
    com_inited = False
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inited = True
    except Exception as e:
        win.q.put(("log", {"msg": f"Warning: COM init failed (pipeline): {e}"}))

    try:
        for key, title_text, module in STEPS:
            count, plain = _step_parts(title_text)
            start_msg = f"Starting: {plain}"
            print(start_msg); win.q.put(("log", {"msg": start_msg}))
            win.q.put(("status", {"key": key, "state": "running", "text": "Running…"}))
            t0 = time.time()

            try:
                runner = _get_runner(module)
            except Exception:
                dt = time.time() - t0
                win.q.put(("status", {"key": key, "state": "fail", "text": f"Failed to import ({dt:.1f}s)"}))
                win.q.put(("log", {"msg": traceback.format_exc()}))
                if stop_on_error:
                    win.q.put(("failed_all", {})); return
                continue

            try:
                runner()
                dt = time.time() - t0
                win.q.put(("status", {"key": key, "state": "done", "text": f"Done ({dt:.1f}s)"}))
                finish_msg = (f"[{count}] Finished" if count else "Finished") + f" ({dt:.1f}s)"
                print(finish_msg); win.q.put(("log", {"msg": finish_msg}))
            except Exception:
                dt = time.time() - t0
                win.q.put(("status", {"key": key, "state": "fail", "text": f"Failed ({dt:.1f}s)"}))
                win.q.put(("log", {"msg": traceback.format_exc()}))
                if stop_on_error:
                    win.q.put(("failed_all", {})); return
                # continue
    finally:
        if com_inited:
            try: pythoncom.CoUninitialize()
            except Exception: pass
        win.q.put(("done_all", {}))


def run_reports(win: MainWindow, selected_reports: list[tuple[str,str,str]]):
    com_inited = False
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inited = True
    except Exception as e:
        win.q.put(("log", {"msg": f"Warning: COM init failed (reports): {e}"}))

    try:
        for key, label, module in selected_reports:
            win.q.put(("log", {"msg": f"Starting report: {label}"}))
            t0 = time.time()
            runner = _get_runner(module)
            runner()
            dt = time.time() - t0
            win.q.put(("log", {"msg": f"Finished report: {label} ({dt:.1f}s)"}))
        win.q.put(("reports_done", {}))
    except Exception:
        win.q.put(("log", {"msg": traceback.format_exc()}))
        win.q.put(("reports_failed", {}))
    finally:
        if com_inited:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

# ------------------ Entry ------------------
def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
