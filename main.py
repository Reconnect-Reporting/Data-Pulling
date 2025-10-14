# main.py (PySide6 version with tabs + responsive flow + Setup tab w/ deps installer)

# -*- coding: utf-8 -*-
# ==== Auto_reports_pulling ====
# ---- Bootstrap: ensure core packages before heavy imports ----
import sys, os, subprocess, importlib.util
from PySide6.QtWidgets import QTableView, QHeaderView
from PySide6.QtCore import QAbstractTableModel, QModelIndex
import pandas as pd
import io
from contextlib import redirect_stdout, redirect_stderr


def _in_venv() -> bool:
    return (hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix) or \
           (hasattr(sys, "real_prefix") and sys.real_prefix != sys.prefix)

# Map "import name" -> "pip package"
_BOOTSTRAP_REQUIRED = {
    "PySide6": "PySide6",
    "pandas": "pandas",
    "numpy": "numpy",
    "openpyxl": "openpyxl",
    "selenium": "selenium",
    "webdriver_manager": "webdriver-manager",
    # pywin32 provides both win32com and pythoncom on Windows
    **({"pythoncom": "pywin32", "win32com": "pywin32"} if os.name == "nt" else {}),
}

def _ensure_base_deps():
    missing = []
    for import_name, pip_name in _BOOTSTRAP_REQUIRED.items():
        if importlib.util.find_spec(import_name) is None:
            missing.append(pip_name)

    if missing:
        # dedupe while keeping deterministic order
        to_install = []
        for p in missing:
            if p not in to_install:
                to_install.append(p)

        cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + to_install
        if not _in_venv():
            cmd.append("--user")

        print("Installing missing packages:", " ".join(to_install))
        # Don't capture output so the user sees progress if it takes a while
        res = subprocess.run(cmd)
        if res.returncode != 0:
            print("Some packages failed to install. Exiting.")
            sys.exit(res.returncode)

        # Re-exec the current process so newly installed modules import cleanly
        os.execv(sys.executable, [sys.executable] + sys.argv)

_ensure_base_deps()


import os, sys, json, time, traceback, importlib, importlib.util, re, threading, queue, subprocess
from PySide6.QtWidgets import QSizePolicy
from pathlib import Path
from typing import Dict, List, Tuple
import importlib, sys
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
    QGraphicsPathItem, QGraphicsDropShadowEffect,QComboBox,QCheckBox,QSplitter
)

# COM (Outlook/Excel) init for worker threads
import pythoncom

# ===================== Dependency Installer (inline) =====================
class DependencyInstaller:
    """Checks & installs third-party deps used by your app (run from Setup tab)."""

    REQUIRED: dict[str, str] = {
        # Core packages
        "pandas": "pandas",
        "numpy": "numpy",
        "openpyxl": "openpyxl",
        "pywin32": "pywin32",                  # replaces win32com.client
        "selenium": "selenium",
        "webdriver_manager": "webdriver-manager",
    
        # Optional but used
        "glob": "glob2",                       # stdlib glob exists, but pip package is glob2 if needed
        "typing_extensions": "typing-extensions",  # for Optional, Iterable, etc. (if backport needed)
    }

    def __init__(self, quiet: bool = True):
        self.quiet = quiet
        self._log: List[str] = []

    def ensure(self) -> Dict[str, str]:
        status: Dict[str, str] = {}
        missing = self._missing()

        if not missing:
            for pip_name in set(self.REQUIRED.values()):
                status[pip_name] = "ok"
            self._add("All dependencies already satisfied ✅")
            return status

        self._add("Missing packages detected:")
        for mod, pip_name in missing:
            self._add(f"  - {mod}  (pip: {pip_name})")

        targets = sorted({pip for _, pip in missing})
        ok, failed = self._pip_install(targets)

        # Re-verify after install
        for import_path, pip_name in self.REQUIRED.items():
            if pip_name in ok:
                status[pip_name] = "installed" if self._can_import(import_path) else "failed: import still missing"
            elif pip_name in failed:
                status[pip_name] = f"failed: {failed[pip_name]}"
            else:
                status[pip_name] = "ok"

        if status.get("pywin32", "").startswith("installed"):
            self._add("Note: If Outlook/Excel automation fails, restart Python/Windows to finalize pywin32.")
        return status

    def get_log(self) -> str:
        return "\n".join(self._log)

    # ---- internals
    def _missing(self) -> List[Tuple[str, str]]:
        out: List[Tuple[str, str]] = []
        for import_path, pip_name in self.REQUIRED.items():
            if not self._can_import(import_path):
                out.append((import_path, pip_name))
        return out

    def _can_import(self, import_path: str) -> bool:
        return importlib.util.find_spec(import_path) is not None

    def _pip_install(self, pkgs: List[str]) -> Tuple[List[str], Dict[str, str]]:
        ok: List[str] = []
        failed: Dict[str, str] = {}

        in_venv = (hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix) or \
                  (hasattr(sys, "real_prefix") and sys.real_prefix != sys.prefix)

        cmd_base = [sys.executable, "-m", "pip", "install", "--upgrade"]
        if not in_venv:
            cmd_base.append("--user")

        for pkg in pkgs:
            cmd = cmd_base + [pkg]
            try:
                self._add(f"Installing {pkg} ...")
                res = subprocess.run(
                    cmd,
                    capture_output=self.quiet,
                    text=True,
                    check=False
                )
                if res.returncode == 0:
                    ok.append(pkg)
                    self._add(f"✓ {pkg} installed/updated")
                else:
                    reason = (res.stderr or res.stdout or "unknown error").strip().splitlines()[-1]
                    failed[pkg] = reason
                    self._add(f"✗ {pkg} failed: {reason}")
            except Exception as e:
                failed[pkg] = str(e)
                self._add(f"✗ {pkg} failed: {e}")
        return ok, failed

    def _add(self, msg: str):
        self._log.append(msg)

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

def _apply_caredove_env(email: str, password: str) -> None:
    os.environ["CAREDOVE_EMAIL"] = email
    os.environ["CAREDOVE_PASSWORD"] = password

def _apply_trapeze_env(username: str, password: str) -> None:
    os.environ["TRAPEZE_USERNAME"] = username
    os.environ["TRAPEZE_PASSWORD"] = password

# ------------------ PIPELINE STEPS ------------------
STEPS = [
    ("clean",  "[1/11] Cleaning Folder",                   "Auto_reports_pulling.Clean_Download_Folder"),
    ("treat",  "[2/11] Pulling Treat Reports",             "Auto_reports_pulling.Treat_Pulling"),
    ("move",   "[3/11] Moving Treat Files to OneDrive",    "Data_Pipeline.Treat_File_Moving"),
    ("alaya",  "[4/11] Pulling AlayaCare Reports",         "Auto_reports_pulling.AlayaCare_Pulling"),
    ("caredove","[5/11] Pulling Caredove Reports",         "Auto_reports_pulling.Caredove_Pulling"),
    ("trapeze","[6/11] Pulling Trapeze Reports",           "Auto_reports_pulling.Trapeze_Report_Pulling"),
    ("cmove",  "[7/11] Moving Caredove Files to OneDrive", "Data_Pipeline.Caredove_File_Moving"),
    ("tclean", "[8/11] Cleaning Treat Data",               "Data_Pipeline.Treat_Data_Cleaning"),
    ("aclean", "[9/11] Cleaning AlayaCare Data",           "Data_Pipeline.AlayaCare_Data_Cleaning"),
    ("cclean", "[10/11] Cleaning Caredove Data",           "Data_Pipeline.Caredove_Data_Cleaning"),
    ("report_dashboard","[11/11] Refreshing Dashboards", [
        "Reports.OHRS_MHA",
        "Reports.OHRS_SFS",
        "Reports.Client_enroll_in_multiple_programs"
    ])
]

# ------------------ REPORTS ------------------
REPORTS = [
    ("hhri", "HHRI (Save to Downloads)", "Reports.HHRI_Hours","weekly"),
    ("fame", "CMHA Peel – FAME Monthly (save to Downloads)", "Reports.FAME_Report","monthly"),
    ("jam",  "JAM Report (Save to Downloads)", "Reports.JAM_Report","monthly"),
    ("ocan","Overdue OCAN List (Save to Downloads)", "Reports.Overdue_OCAN_List","monthly"),
    ("cwcd","CWCD (show in Output Table)", "Reports.CWCD","monthly"),
    ("uhn",  "UHN Geri Psych (show in Output Table)", "Reports.UHN_GeriPsych","monthly"),
    ("rass",  "RASS (show in Output Table)", "Reports.RASS","monthly"),
    ("ohrs", "OHRS_MHA (show in Output Table and Refresh Dashboard)", "Reports.OHRS_MHA","daily"),
    ("ohrs_sfs", "OHRS_SFS (show in Output Table and Refresh Dashboard)", "Reports.OHRS_SFS","daily"),
    ("consent_missing", "List of Outstanding consents (Save to Downloads)", "Reports.List_of_outstanding_consent","monthly"),
    ("client_in_multiple_programs","Client Enroll in Multiple Programs (show in Output Table and Refresh Dashboard)","Reports.Client_enroll_in_multiple_programs","daily"),
    ("vhs", "VHA (show in Output Table)", "Reports.VHA","monthly"),
    ("geri hub", "Geri Hub Report (show in Output Table)", "Reports.Geri_Hub_Report","monthly"),
    ("ipc", "IPC Report (show in Output Table)", "Reports.IPC_Report","monthly"),
    ("coss", "COSS Report (show in Output Table)", "Reports.COSS_Report","monthly")
]

STATUS_COLORS = {
    "pending": "#666666",
    "running": "#1f6feb",
    "done":    "#18794e",
    "fail":    "#b54708",
}

# ---------- Visual labels for the flow nodes (not the row titles) ----------
FLOW_LABELS = {
    "clean": "Clean Folder",
    "treat_box": "Treat",
    "raw": "Raw Data",
    "alayacare": "AlayaCare",
    "clean_data": "Clean Data",
    "downloads": "Downloads",
    "email": "Email Inbox",
    "caredove": "CareDove",
    "caredove_dl": "Downloads",
    "Trapeze": "Trapeze",        # <-- use 'Trapeze' consistently
    "trapeze_dl": "Downloads",
    "PowerBI": "Power BI",
    "reports": "Reports ready",
}


# Map pipeline step -> which node(s) should light up
STEP_NODE_MAP = {
    "clean":  ["clean"],
    "treat":  ["treat_box","downloads"],
    "move":   [],
    "alaya":  ["alayacare","email","raw"],
    "tclean": [],
    "aclean": [],
    "caredove":["caredove","caredove_dl"],
    "cmove":["raw"],
    "cclean":["clean_data"],
    "trapeze": ["Trapeze", "trapeze_dl"], 
    "report_dashboard": ["PowerBI", "reports"],

}

# Which edges should change per pipeline step
STEP_EDGE_MAP = {
    # after CLEAN finishes, light up all four outgoing edges from 'clean'
    "clean":  [("clean","treat_box"),
               ("clean","alayacare"),
               ("clean","caredove"),
               ("clean","Trapeze")],

    "treat":  [("treat_box", "downloads")],
    "move":   [("downloads", "raw")],
    "alaya":  [("alayacare", "email"), ("email","raw")],
    "tclean": [],
    "aclean": [("raw", "clean_data")],
    "caredove":[("caredove","caredove_dl")],
    "cmove":  [("caredove_dl","raw")],
    "cclean":[("raw", "clean_data")],
    # keep node key consistent: use 'Trapeze' everywhere
    "trapeze": [("Trapeze","trapeze_dl"), ("trapeze_dl","raw")],
    "report_dashboard": [("clean_data","PowerBI"), ("clean_data","reports")]
}

# All edges in the diagram (used for resetting)
ALL_EDGES = [
    ("clean","treat_box"),
    ("clean","alayacare"),
    ("clean","caredove"),
    ("clean","Trapeze"),
    ("treat_box","downloads"),
    ("downloads","raw"),
    ("alayacare","email"),
    ("email","raw"),
    ("raw","clean_data"),
    ("caredove","caredove_dl"),
    ("caredove_dl","raw"),
    ("Trapeze","trapeze_dl"),
    ("trapeze_dl","raw"),
    ("clean_data","PowerBI"),
    ("clean_data","reports"),
]

# ------------------ Helpers ------------------
def _step_parts(title_text: str):
    m = re.match(r"^\[(\d+/\d+)\]\s*(.+)$", title_text.strip())
    return (m.group(1), m.group(2)) if m else (None, title_text)

def _get_runner(module_spec):
    """
    Accepts either:
      - str: 'pkg.module' that exposes run() or main()
      - list/tuple[str]: several modules; returns a wrapper that runs each in order
    """
    def _import_runner(modname: str):
        if modname in sys.modules:
            mod = importlib.reload(sys.modules[modname])
        else:
            mod = importlib.import_module(modname)
        fn = getattr(mod, "run", None) or getattr(mod, "main", None)
        if not callable(fn):
            raise AttributeError(f"{modname} has no callable run() or main().")
        return fn

    # multiple modules
    if isinstance(module_spec, (list, tuple)):
        fns = [_import_runner(m) for m in module_spec]

        def run_all():
            for fn in fns:
                fn()
        return run_all

    # single module
    if isinstance(module_spec, str):
        return _import_runner(module_spec)

    raise TypeError(f"Unsupported module spec type: {type(module_spec).__name__}")


# ------------------ Qt FlowView with optional picture nodes ------------------
# ------------------ Qt FlowView with optional picture nodes ------------------
class FlowView(QGraphicsView):

    # --- Responsive layout baseline (used only as a fallback) ---
    DESIGN_W = 4000
    DESIGN_H = 600

    DEFAULT_ICON_SIZE = 220
    BASE_ICON_SIZES = {
        "clean": 220, "treat_box": 400, "alayacare": 400,
        "downloads": 180, "email": 180, "raw": 220, "clean_data": 220,
        "caredove_dl": 180, "caredove": 450, "Trapeze": 400, "trapeze_dl": 180,
        "PowerBI": 400, "reports": 400,
    }

    BASE_NODE_POS = {
        "clean": (180, 335),
        "treat_box": (900, 10),
        "alayacare": (900, 220),
        "downloads": (1500, 10),
        "email": (1500, 220),
        "raw": (1950, 335),
        "clean_data": (2400, 335),
        "caredove": (900, 450),
        "caredove_dl": (1500, 450),
        "Trapeze": (900, 650),
        "trapeze_dl": (1500, 650),
        "PowerBI": (2900, 150),
        "reports": (2900, 520),
    }

    SCENE_PADDING = 60  # extra breathing room

    def __init__(self, parent=None, icons_dir: Path | None = None):
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self.setScene(self._scene)
        self.setRenderHint(QPainter.Antialiasing, True)
        self.setRenderHint(QPainter.TextAntialiasing, True)
        self.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.setFrameShape(QFrame.NoFrame)
        self.setMinimumHeight(300)

        self.icons_dir = icons_dir or (Path(__file__).resolve().parent / "icons")

        # state storage
        self._nodes: dict = {}   # key -> {"pix": item, "bbox": (x0,y0,x1,y1), "effect": effect}
        self._edges: list = []   # list of dicts per edge
        self._states = {
            "clean": "pending", "treat_box": "pending", "alayacare": "pending",
            "downloads": "pending", "email": "pending", "raw": "pending", "clean_data": "pending",
        }

        # compute dynamic design bounds before first rebuild
        self._ensure_design_bounds()
        self.rebuild()

    # ---------- bounds helpers ----------
    def _compute_design_bounds(self) -> tuple[float, float]:
        minx = miny = float("inf")
        maxx = maxy = float("-inf")
        for key, (cx, cy) in self.BASE_NODE_POS.items():
            sz = self.BASE_ICON_SIZES.get(key, self.DEFAULT_ICON_SIZE)
            half = sz / 2.0
            minx = min(minx, cx - half)
            maxx = max(maxx, cx + half)
            miny = min(miny, cy - half)
            maxy = max(maxy, cy + half)
        w = max(1.0, maxx - minx)
        h = max(1.0, maxy - miny)
        return w, h

    def _ensure_design_bounds(self) -> None:
        if not hasattr(self, "_design_w") or not hasattr(self, "_design_h"):
            self._design_w, self._design_h = self._compute_design_bounds()

    # ---------- responsive helpers ----------
    def _scale_factor(self) -> float:
        self._ensure_design_bounds()
        vw = max(1, self.viewport().width()  - 2*self.SCENE_PADDING)
        vh = max(1, self.viewport().height() - 2*self.SCENE_PADDING)
        sx = vw / self._design_w
        sy = vh / self._design_h
        s = min(sx, sy)
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

        # Recompute dynamic design bounds each time (in case layout changed)
        self._design_w, self._design_h = self._compute_design_bounds()
        s = self._scale_factor()

        # Scene rect sized to dynamic design, with padding
        pad = self.SCENE_PADDING
        scene_w = self._design_w * s
        scene_h = self._design_h * s
        self._scene.setSceneRect(-pad, -pad, scene_w + 2*pad, scene_h + 2*pad)

        # place nodes using scaled positions & sizes
        for key in self.BASE_NODE_POS.keys():
            cx, cy = self._scaled_pos(key, s)
            size_px = self._scaled_icon_size(key, s)
            self._add_node_image(key, cx, cy, size_px)

        # arrows (same logical flow)
        self._add_arrow("clean", "treat_box")
        self._add_arrow("clean", "alayacare")
        self._add_arrow("clean", "caredove")
        self._add_arrow("clean", "Trapeze")
        self._add_arrow("treat_box", "downloads")
        self._add_arrow("alayacare", "email")
        self._add_arrow("downloads", "raw")
        self._add_arrow("email", "raw")
        self._add_arrow("raw", "clean_data")
        self._add_arrow("caredove", "caredove_dl")
        self._add_arrow("caredove_dl", "raw")
        self._add_arrow("Trapeze", "trapeze_dl")
        self._add_arrow("trapeze_dl", "raw")
        self._add_arrow("clean_data", "PowerBI")
        self._add_arrow("clean_data", "reports")

        # re-apply saved states (steady glow)
        for k, st in self._states.items():
            self._apply_state(k, st)

        self._edge_states: dict[tuple[str, str], str] = {}

    def resizeEvent(self, e):
        super().resizeEvent(e)
        # Keep full rebuild (you scale positions & sizes yourself)
        QTimer.singleShot(0, self.rebuild)

    # ---------- node state (steady glow) ----------
    def _state_color(self, state: str) -> QColor:
        if state in ("running", "done"): return QColor("#18794e")   # GREEN
        if state == "fail":              return QColor("#b54708")
        return QColor(0, 0, 0, 0)

    def _apply_state(self, key, state):
        self._states[key] = state
        node = self._nodes.get(key)
        if not node:
            return
        eff: QGraphicsDropShadowEffect = node["effect"]
        col = self._state_color(state)
        if state == "pending":
            eff.setBlurRadius(0); eff.setColor(QColor(0, 0, 0, 0))
        else:
            eff.setBlurRadius(28); eff.setColor(col)

    # external API used by MainWindow
    def set_node_state(self, key, state):
        self._apply_state(key, state)


class PandasModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame, parent=None):
        super().__init__(parent)
        self._df = df.copy()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role):
        if not index.isValid() or role not in (Qt.DisplayRole, Qt.EditRole):
            return None

        val = self._df.iat[index.row(), index.column()]

        # 🔒 Hide empties: NaN/None/NA or the literal string "nan"
        if pd.isna(val) or (isinstance(val, str) and val.strip().lower() == "nan"):
            return ""

        # nice display for floats/ints
        if isinstance(val, float):
            try:
                return f"{val:.0f}" if float(val).is_integer() else f"{val:.2f}"
            except Exception:
                return ""  # just in case val was a weird float-like

        return str(val)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(self._df.index[section])


# ------------------ Main Window ------------------
class MainWindow(QMainWindow):
    def _reset_visuals(self):
        # reset rows
        for k, r in self.step_rows.items():
            r["status"].setText("Pending")
            r["status"].setStyleSheet(f"color:{STATUS_COLORS['pending']};")
        # reset nodes
        for node_key in FLOW_LABELS:
            self.flow.set_node_state(node_key, "pending")
        # reset edges
        for a, b in ALL_EDGES:
            self.flow.set_edge_state(a, b, "pending")


    def start_checked_steps(self):
        if self.running:
            return
        # collect the subset in pipeline order
        subset = [(k,t,m) for (k,t,m) in STEPS if self.step_checks.get(k) and self.step_checks[k].isChecked()]
        if not subset:
            self._log("No steps checked."); 
            return

        self._reset_visuals()
        self.running = True
        self.btn_run.setEnabled(False); self.btn_run.setText("Running…")
        self.btn_run_checked.setEnabled(False); self.btn_run_checked.setText("Running…")
        threading.Thread(target=run_pipeline, args=(self, subset, True), daemon=True).start()
    
    def _show_dataframe(self, df: pd.DataFrame):
        self.out_table.setModel(PandasModel(df))
        # Jump to the Output tab so the user sees the result immediately
        if hasattr(self, "tab_out"):
            self.tabs.setCurrentWidget(self.tab_out)

    
    def preview_cwcd_pivot(self):
        try:
            import Reports.CWCD as cwcd
            df = cwcd.run()  # uses your existing transformation
            self._show_dataframe(df)
            self._log("CWCD pivot built and rendered in Output Table.")
        except Exception as e:
            self._log(f"Error building CWCD pivot: {e}")


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
        if "CAREDOVE_EMAIL" in self.settings:
            os.environ["CAREDOVE_EMAIL"] = self.settings["CAREDOVE_EMAIL"]
        if "CAREDOVE_PASSWORD" in self.settings:
            os.environ["CAREDOVE_PASSWORD"] = self.settings["CAREDOVE_PASSWORD"]
        if "TRAPEZE_USERNAME" in self.settings:
            os.environ["TRAPEZE_USERNAME"] = self.settings["TRAPEZE_USERNAME"]
        if "TRAPEZE_PASSWORD" in self.settings:
            os.environ["TRAPEZE_PASSWORD"] = self.settings["TRAPEZE_PASSWORD"]

        # thread queue
        self.q = queue.Queue()
        self.running = False

        # --- central layout ---
        central = QWidget(self)
        self.setCentralWidget(central)
        v = QVBoxLayout(central)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(8)

        self.tabs = QTabWidget()
        v.addWidget(self.tabs, 1)

        self.tab_auto  = QWidget()
        self.tab_rep   = QWidget()
        self.tab_out   = QWidget()   # NEW: Output tab

        self.tabs.addTab(self.tab_auto,  "Data Automation")
        self.tabs.addTab(self.tab_rep,   "Reports")
        self.tabs.addTab(self.tab_out,   "Output")   # NEW


        # ---- Data Automation tab ----
        auto_layout = QVBoxLayout(self.tab_auto)
        auto_layout.setSpacing(8)

        # Header with Run
        hdr = QHBoxLayout()

        back_lbl = QLabel("Data Automation")
        back_lbl.setStyleSheet("font-size:18px; font-weight:600;")
        hdr.addWidget(back_lbl)
        hdr.addStretch(1)

        self.btn_run_checked = QPushButton("Run checked")
        self.btn_run_checked.clicked.connect(self.start_checked_steps)
        hdr.addWidget(self.btn_run_checked)

        self.btn_run = QPushButton("Run")
        self.btn_run.clicked.connect(self.start_pipeline)
        hdr.addWidget(self.btn_run)

        # 🔽 Toggle Treat credentials
        self.btn_toggle_creds = QPushButton("Update Credentials")
        self.btn_toggle_creds.clicked.connect(self._toggle_creds)
        hdr.addWidget(self.btn_toggle_creds)

        auto_layout.addLayout(hdr)

        # Credentials (Treat + Caredove) - toggled by the button
        self.creds_group = QGroupBox("Credentials")
        gl = QGridLayout(self.creds_group)
        gl.setColumnStretch(1, 1)
        gl.setColumnStretch(3, 1)

        # --- Treat ---
        treat_title = QLabel("<b>Treat</b>")
        gl.addWidget(treat_title, 0, 0, 1, 5)

        self.in_user = QLineEdit(os.getenv("TREAT_USERNAME", self.settings.get("TREAT_USERNAME", "yxu")))
        self.in_pass = QLineEdit(os.getenv("TREAT_PASSWORD", self.settings.get("TREAT_PASSWORD", "")))
        self.in_pass.setEchoMode(QLineEdit.Password)

        gl.addWidget(QLabel("Username"), 1, 0); gl.addWidget(self.in_user, 1, 1)
        gl.addWidget(QLabel("Password"), 1, 2); gl.addWidget(self.in_pass, 1, 3)



        # --- Caredove ---
        caredove_title = QLabel("<b>Caredove</b>")
        gl.addWidget(caredove_title, 3, 0, 1, 5)

        self.cd_email = QLineEdit(os.getenv("CAREDOVE_EMAIL", self.settings.get("CAREDOVE_EMAIL", "")))
        self.cd_pass  = QLineEdit(os.getenv("CAREDOVE_PASSWORD", self.settings.get("CAREDOVE_PASSWORD", "")))
        self.cd_pass.setEchoMode(QLineEdit.Password)

        gl.addWidget(QLabel("Email"),    4, 0); gl.addWidget(self.cd_email, 4, 1)
        gl.addWidget(QLabel("Password"), 4, 2); gl.addWidget(self.cd_pass,  4, 3)

        # --- Trapeze ---
        trapeze_title = QLabel("<b>Trapeze</b>")
        gl.addWidget(trapeze_title, 6, 0, 1, 5)
        
        self.tr_user = QLineEdit(os.getenv("TRAPEZE_USERNAME", self.settings.get("TRAPEZE_USERNAME", "")))
        self.tr_pass = QLineEdit(os.getenv("TRAPEZE_PASSWORD", self.settings.get("TRAPEZE_PASSWORD", "")))
        self.tr_pass.setEchoMode(QLineEdit.Password)
        
        gl.addWidget(QLabel("Username"), 7, 0); gl.addWidget(self.tr_user, 7, 1)
        gl.addWidget(QLabel("Password"), 7, 2); gl.addWidget(self.tr_pass, 7, 3)
        
        # Move Save button BELOW Trapeze (row 8), span all columns, right-aligned
        self.btn_save_creds = QPushButton("Save")
        self.btn_save_creds.clicked.connect(self._save_creds)
        gl.addWidget(self.btn_save_creds, 8, 0, 1, 5, alignment=Qt.AlignRight)



        # Flow (with optional icons)
        self.flow = FlowView(self, icons_dir=Path(__file__).resolve().parent / "icons")
        auto_layout.addWidget(self.flow)

        # Scrollable steps list
        # ---------- Scrollable steps list (drop-in replacement) ----------
        steps_group = QGroupBox("Pipeline steps")
        steps_v = QVBoxLayout(steps_group)

        self.steps_scroll = QScrollArea()
        self.steps_scroll.setWidgetResizable(True)

        steps_widget = QWidget()
        self.steps_scroll.setWidget(steps_widget)

        # Create the layout BEFORE using it anywhere
        self.steps_layout = QVBoxLayout(steps_widget)
        self.steps_layout.setContentsMargins(8, 8, 8, 8)
        self.steps_layout.setSpacing(4)

        # keep rows + add per-step checkboxes
        self.step_rows = {}
        self.step_checks = {}  # key -> QCheckBox

        for key, title_text, module in STEPS:
            row = QWidget()
            hl = QHBoxLayout(row)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.setSpacing(8)

            # per-step checkbox
            cb = QCheckBox()
            cb.setObjectName(f"cb_{key}")
            self.step_checks[key] = cb
            hl.addWidget(cb, 0)

            lbl = QLabel(title_text)
            lbl.setMinimumWidth(600)

            # Accept str or list/tuple[str]
            if isinstance(module, (list, tuple)):
                tip = "\n".join(map(str, module))
            else:
                tip = str(module)
            lbl.setToolTip(tip)

            st = QLabel("Pending")
            st.setStyleSheet(f"color:{STATUS_COLORS['pending']};")

            hl.addWidget(lbl, 1)
            hl.addStretch(1)
            hl.addWidget(st, 0)

            self.steps_layout.addWidget(row)
            self.step_rows[key] = {"row": row, "label": lbl, "status": st, "checkbox": cb}

        # push items to top
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.steps_layout.addWidget(spacer)

        steps_v.addWidget(self.steps_scroll)
        auto_layout.addWidget(steps_group)



        # Controls
        ctl = QHBoxLayout()
        self.progress = QProgressBar(); self.progress.setRange(0,0); self.progress.setVisible(False)
        ctl.addWidget(self.progress); ctl.addStretch(1)
        auto_layout.addLayout(ctl)

        # ---- Output tab (full-bleed table) ----
        out_layout = QVBoxLayout(self.tab_out)
        out_layout.setContentsMargins(0, 0, 0, 0)
        out_layout.setSpacing(0)
        
        self.out_table = QTableView(self.tab_out)
        self.out_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.out_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.out_table.horizontalHeader().setStretchLastSection(True)
        self.out_table.verticalHeader().setVisible(False)
        self.out_table.setSortingEnabled(True)
        self.out_table.setAlternatingRowColors(True)
        self.out_table.setSelectionBehavior(QTableView.SelectRows)
        self.out_table.setSelectionMode(QTableView.SingleSelection)
        self.out_table.setWordWrap(False)
        self.out_table.setCornerButtonEnabled(False)
        out_layout.addWidget(self.out_table)
        
        # ---- Reports tab ----
        rep_layout = QVBoxLayout(self.tab_rep)
        rep_layout.setSpacing(8)
        rep_layout.addWidget(QLabel("<b>Reports</b>"))
        
        # Split the page vertically into Daily / Weekly / Monthly
        splitter = QSplitter(Qt.Vertical, self.tab_rep)
        rep_layout.addWidget(splitter, 1)
        
        self.report_checks = {}     # key -> checkbox (define ONCE)
        self._section_checks = {}   # cadence -> list[checkboxes]
        
        def build_section(title: str, cadence_key: str):
            box = QGroupBox(title)
            v = QVBoxLayout(box); v.setSpacing(6)
        
            # Select-all for this section
            row = QWidget(); hl = QHBoxLayout(row); hl.setContentsMargins(0,0,0,0)
            sel_all = QCheckBox("Select all in this section")
            hl.addWidget(sel_all); hl.addStretch(1)
            v.addWidget(row)
        
            # Container for the report checkboxes
            container = QWidget(); cv = QVBoxLayout(container)
            cv.setContentsMargins(0,0,0,0); cv.setSpacing(4)
            v.addWidget(container, 1)
        
            checks_in_section = []
            # REPORTS is now (key, label, module, cadence)
            for key, label, module, cadence in REPORTS:
                if cadence != cadence_key:
                    continue
                cb = QCheckBox(label)
                cb.setObjectName(f"rep_{key}")
                cb.setToolTip(str(module))
                cv.addWidget(cb)
                self.report_checks[key] = cb
                checks_in_section.append(cb)
        
            # Select-all behavior for this section
            def toggle_all(state):
                on = state == Qt.Checked
                for c in checks_in_section:
                    c.setChecked(on)
            sel_all.stateChanged.connect(toggle_all)
        
            self._section_checks[cadence_key] = checks_in_section
            return box
        
        daily_box   = build_section("Daily",   "daily")
        weekly_box  = build_section("Weekly",  "weekly")
        monthly_box = build_section("Monthly", "monthly")
        
        splitter.addWidget(daily_box)
        splitter.addWidget(weekly_box)
        splitter.addWidget(monthly_box)
        
        # Optional: nicer initial splitter sizes
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 1)
        
        # Run button at the bottom (ONE instance)
        self.btn_run_reports = QPushButton("Create report(s)")
        self.btn_run_reports.clicked.connect(self.start_reports)
        rep_layout.addWidget(self.btn_run_reports)


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
       # Treat
       u = self.in_user.text().strip()
       p = self.in_pass.text()

       # Caredove
       cd_email = self.cd_email.text().strip()
       cd_pass  = self.cd_pass.text()

       # Trapeze
       tr_u = self.tr_user.text().strip()
       tr_p = self.tr_pass.text()

       # Apply to env immediately
       _apply_treat_env(u, p)
       _apply_caredove_env(cd_email, cd_pass)
       _apply_trapeze_env(tr_u, tr_p)

       # Persist
       s = dict(_load_settings())
       s["TREAT_USERNAME"]     = u
       s["TREAT_PASSWORD"]     = p
       s["CAREDOVE_EMAIL"]     = cd_email
       s["CAREDOVE_PASSWORD"]  = cd_pass
       s["TRAPEZE_USERNAME"]   = tr_u
       s["TRAPEZE_PASSWORD"]   = tr_p
       _save_settings(s)

       self._log("Saved Treat, Caredove & Trapeze credentials (env updated).")



    # ---- logging
    def _log(self, msg: str):
        self.log.appendPlainText(msg.rstrip())
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def _log_setup(self, msg: str):
        self.setup_log.appendPlainText(msg.rstrip())
        self.setup_log.verticalScrollBar().setValue(self.setup_log.verticalScrollBar().maximum())

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
                    self.btn_run_checked.setEnabled(True); self.btn_run_checked.setText("Run checked")

                elif typ == "failed_all":
                    self._set_progress(False)
                    self._log("Pipeline stopped due to failure.")
                    self.running = False
                    self.btn_run.setEnabled(True); self.btn_run.setText("Run")
                    self.btn_run_checked.setEnabled(True); self.btn_run_checked.setText("Run checked")
                elif typ == "reports_done":
                    self._log("Report job(s) completed.")
                    self.btn_run_reports.setEnabled(True); self.btn_run_reports.setText("Create report(s)")
                elif typ == "reports_failed":
                    self._log("Report job(s) stopped due to failure.")
                    self.btn_run_reports.setEnabled(True); self.btn_run_reports.setText("Create report(s)")
                elif typ == "deps_log":
                    self._log_setup(payload["msg"])
                elif typ == "deps_done":
                    self.btn_install_deps.setEnabled(True); self.btn_install_deps.setText("Check & Install Libraries")
                    self._log_setup("Dependency check finished.")
                elif typ == "deps_failed":
                    self.btn_install_deps.setEnabled(True); self.btn_install_deps.setText("Check & Install Libraries")
                    self._log_setup("Dependency check stopped due to failure.")
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
        for a, b in STEP_EDGE_MAP.get(step_key, []):
            # Only upgrade state, don’t downgrade
            prev_state = next((e["state"] for e in self.flow._edges if e["from"] == a and e["to"] == b), "pending")
            if prev_state != "done" or state == "fail":
                self.flow.set_edge_state(a, b, state)


    # ---- Actions
    # inside MainWindow
    def _reset_visuals(self):
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

    def start_pipeline(self):
        if self.running: return
        self._reset_visuals()
        self.running = True
        self.btn_run.setEnabled(False); self.btn_run.setText("Running…")
        # runs all steps because steps=None by default
        threading.Thread(target=run_pipeline, args=(self,), daemon=True).start()


    def start_reports(self):
        selected = [(key, label, module) for (key, label, module, cadence) in REPORTS
                if self.report_checks[key].isChecked()]
        if not selected:
            self._log("No reports selected."); return
        self.btn_run_reports.setEnabled(False); self.btn_run_reports.setText("Running…")
        threading.Thread(target=run_reports, args=(self, selected), daemon=True).start()

    def start_deps_check(self):
        self.btn_install_deps.setEnabled(False)
        self.btn_install_deps.setText("Running…")
        self.setup_log.clear()
        threading.Thread(target=run_deps_check, args=(self,), daemon=True).start()

# ------------------ Workers ------------------
def run_pipeline(win: "MainWindow", steps: list[tuple[str,str,str]] | None = None, stop_on_error: bool = True):
    if steps is None:
        steps = STEPS  # default: all

    com_inited = False
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inited = True
    except Exception as e:
        win.q.put(("log", {"msg": f"Warning: COM init failed (pipeline): {e}"}))

    try:
        for key, title_text, module in steps:
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
    finally:
        if com_inited:
            try: pythoncom.CoUninitialize()
            except Exception: pass
        win.q.put(("done_all", {}))

# --- QUIET reports: only a single success/fail line per report ---
def run_reports(win: MainWindow, selected_reports: list[tuple[str,str,str]]):
    com_inited = False
    try:
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            com_inited = True
        except Exception as e:
            win.q.put(("log", {"msg": f"COM init warning: {e}"}))

        for key, label, module in selected_reports:
            try:
                runner = _get_runner(module)
                result = runner()  # Module decides what to return

                # Show in Output table if we get a DataFrame,
                # OR if we get a path to a CSV/XLSX file.
                df_to_show = None

                if isinstance(result, pd.DataFrame):
                    df_to_show = result

                elif isinstance(result, (str, Path)):
                    p = Path(result)
                    if p.suffix.lower() == ".csv":
                        # Your CSS pipeline writes CSV w/o headers; show as-is
                        df_to_show = pd.read_csv(p, header=None, dtype="object", encoding="utf-8-sig")
                    elif p.suffix.lower() in (".xlsx", ".xls"):
                        # If a report writes Excel, preview it too
                        df_to_show = pd.read_excel(p, header=None, dtype="object")

                if df_to_show is not None:
                    win._show_dataframe(df_to_show)

                # ONE LINE ONLY ✅
                if isinstance(result, (str, Path)):
                    win.q.put(("log", {"msg": f"{label} — file updated: {Path(result)}"}))
                else:
                    win.q.put(("log", {"msg": f"{label} — successfully generated."}))

            except Exception as e:
                win.q.put(("log", {"msg": f"{label} — failed: {e}"}))
                win.q.put(("reports_failed", {}))
                return

        win.q.put(("reports_done", {}))
    finally:
        if com_inited:
            try: pythoncom.CoUninitialize()
            except Exception: pass




def run_deps_check(win: MainWindow):
    try:
        win.q.put(("deps_log", {"msg": "Starting dependency check..."}))
        installer = DependencyInstaller(quiet=True)
        status = installer.ensure()
        log = installer.get_log()
        if log.strip():
            win.q.put(("deps_log", {"msg": log}))

        lines = ["", "Status summary:"]
        for pkg, st in status.items():
            lines.append(f"  {pkg:10s} -> {st}")
        win.q.put(("deps_log", {"msg": "\n".join(lines)}))

        problems = [k for k, v in status.items() if v.startswith("failed")]
        if problems:
            win.q.put(("deps_log", {"msg": "\nSome packages failed to install:\n- " + "\n- ".join(f"{p}: {status[p]}" for p in problems)}))
        win.q.put(("deps_done", {}))
    except Exception:
        win.q.put(("deps_log", {"msg": traceback.format_exc()}))
        win.q.put(("deps_failed", {}))

# ------------------ Entry ------------------
def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
