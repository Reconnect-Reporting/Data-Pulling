# main.py
import importlib
import threading
import time
import traceback
import queue
import sys
import os
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk
import re
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

# ------------------ PIPELINE STEPS (execution order unchanged) ------------------
STEPS = [
    ("clean",  "[1/6] Cleaning Folder",                 "Clean_Download_Folder"),
    ("treat",  "[2/6] Pulling Treat Reports",           "Treat_Pulling"),
    ("move",   "[3/6] Move Treat files to One-Drive",   "Treat_File_Moving"),
    ("tclean", "[4/6] Cleaning Treat Reports",          "Treat_Data_Cleaning"),
    ("alaya",  "[5/6] Pulling AlayaCare Reports",       "AlayaCare_Pulling"),
    ("aclean", "[6/6] AlayaCare Data Cleaning",         "AlayaCare_Data_Cleaning"),
]

# ------------------ REPORTS ------------------
REPORTS = [
    ("hhri", "HHRI (email password-protected Excel)", "HHRI_Hours"),
    ("fame", "CMHA Peel – FAME Monthly (email Excel)", "FAME_Report"),
    ("jam",  "JAM Report (email Excel)", "JAM_Report"),
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
    "alayacare":    "AlayaCare",
    "downloads":    "Downloads",
    "email":        "Email Inbox",   # below Downloads
    "raw":          "Raw Data",
    "clean_data":   "Clean Data",
}

# ---------- Map pipeline step -> which node(s) should light up ----------
STEP_NODE_MAP = {
    "clean":  ["clean"],
    "treat":  ["treat_box", "downloads"],
    "move":   ["raw"],
    "tclean": ["clean_data"],
    "alaya":  ["alayacare", "email"],
    "aclean": ["clean_data"],
}

# ------------------ Helpers ------------------

def ensure_shortcuts():
    try:
        import win32com.client  # pip install pywin32
    except ImportError:
        return  # if pywin32 isn't installed, just skip

    from pathlib import Path
    import os

    root = Path(__file__).resolve().parent
    target = str(root / "DailyAutomation.bat")  # shortcut launches your batch
    icon   = str(root / "app.ico")             # include app.ico in repo root

    shell = win32com.client.Dispatch("WScript.Shell")

    # Desktop
    desktop_dir = Path(os.environ["USERPROFILE"]) / "Desktop"
    desktop_dir.mkdir(parents=True, exist_ok=True)
    desktop_lnk = desktop_dir / "Daily Automation.lnk"

    # Start Menu
    start_menu_dir = Path(os.environ["APPDATA"]) / r"Microsoft\Windows\Start Menu\Programs"
    start_menu_dir.mkdir(parents=True, exist_ok=True)
    start_lnk = start_menu_dir / "Daily Automation.lnk"

    for link_path in (desktop_lnk, start_lnk):
        if not link_path.exists():
            sc = shell.CreateShortcut(str(link_path))
            sc.TargetPath = target
            sc.WorkingDirectory = str(root)
            sc.IconLocation = icon if os.path.exists(icon) else target
            sc.Description = "Daily Automation"
            sc.Save()


def _step_parts(title_text: str):
    m = re.match(r"^\[(\d+/\d+)\]\s*(.+)$", title_text.strip())
    return (m.group(1), m.group(2)) if m else (None, title_text)

def _get_runner(module_name: str):
    # Import or reload so modules re-read env vars each step
    if module_name in sys.modules:
        mod = importlib.reload(sys.modules[module_name])
    else:
        mod = importlib.import_module(module_name)
    if hasattr(mod, "run") and callable(mod.run):
        return mod.run
    if hasattr(mod, "main") and callable(mod.main):
        return mod.main
    raise AttributeError(
        f"{module_name} has no callable run() or main(). "
        "Wrap your script logic in a function named run() or main()."
    )

# ------------------ A simple reusable scrollable frame ------------------
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, height=200):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.inner = ttk.Frame(self.canvas)

        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        # set a fixed visible height; contents can be taller and will scroll
        self.canvas.configure(height=height)

        # smooth mouse-wheel scrolling while cursor is over the widget
        self.inner.bind("<Enter>", lambda e: self._bind_mousewheel(True))
        self.inner.bind("<Leave>", lambda e: self._bind_mousewheel(False))

    def _on_inner_configure(self, event):
        # Update scrollregion to encompass the inner frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # Make inner frame width match canvas width
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _bind_mousewheel(self, bind):
        func = self._on_mousewheel
        widget = self.canvas
        if bind:
            widget.bind_all("<MouseWheel>", func)      # Windows
            widget.bind_all("<Button-4>", func)        # Linux up
            widget.bind_all("<Button-5>", func)        # Linux down
        else:
            widget.unbind_all("<MouseWheel>")
            widget.unbind_all("<Button-4>")
            widget.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        # Normalize wheel delta across platforms
        if event.num == 4 or getattr(event, "delta", 0) > 0:
            self.canvas.yview_scroll(-1, "units")
        else:
            self.canvas.yview_scroll(1, "units")

# ============================== App ==============================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Daily Automation")
        self.geometry("1100x760")  # a bit taller (bigger log box)

        try: self.iconbitmap("app.ico")
        except Exception: pass
        try: ttk.Style().theme_use("vista")
        except Exception: pass

        style = ttk.Style()
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Big.TButton", font=("Segoe UI", 14, "bold"), padding=(20, 14))
        style.configure("Back.TButton", font=("Segoe UI", 10), padding=(12, 8))

        # Load saved creds (if any) and apply to env so steps import with them
        self.settings = _load_settings()
        if "TREAT_USERNAME" in self.settings:
            os.environ["TREAT_USERNAME"] = self.settings["TREAT_USERNAME"]
        if "TREAT_PASSWORD" in self.settings:
            os.environ["TREAT_PASSWORD"] = self.settings["TREAT_PASSWORD"]

        self.q = queue.Queue()
        self.worker: threading.Thread | None = None
        self.running = False

        self.content = ttk.Frame(self, padding=16)
        self.content.pack(fill="both", expand=True)

        self.log_frame = ttk.Frame(self, padding=(16, 0, 16, 16))
        self.log_frame.pack(fill="both", expand=False)
        self.log = tk.Text(self.log_frame, height=14, wrap="word")  # larger box
        self.log.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(self.log_frame, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.configure(state="disabled")
        self.log["yscrollcommand"] = sb.set

        self.frames = {}
        self._build_frames()
        self.show("menu")

        self.after(100, self._drain_queue)

        self.bind("<Escape>", lambda e: self.show("menu"))
        self.bind_all("<F5>", lambda e: self.frames["automation"].start_pipeline())
        self.bind_all("<Control-r>", lambda e: self.frames["automation"].start_pipeline())

    def _build_frames(self):
        self.frames["menu"] = MenuScreen(self.content, self)
        self.frames["automation"] = AutomationScreen(self.content, self)
        self.frames["reports"] = ReportsScreen(self.content, self)
        for f in self.frames.values():
            f.place(relx=0, rely=0, relwidth=1, relheight=1)

    def show(self, name: str):
        for _, f in self.frames.items():
            f.lower()
        self.frames[name].lift()
        if name == "automation":
            self.frames[name].reset_status()
        self._log(f"View: {name.capitalize()}")

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _drain_queue(self):
        try:
            while True:
                typ, payload = self.q.get_nowait()
                if typ == "status":
                    self.frames["automation"].set_status(payload["key"], payload["state"], payload.get("text"))
                    if payload["state"] == "running":
                        self.frames["automation"].set_progress(True)
                elif typ == "log":
                    self._log(payload["msg"])
                elif typ == "done_all":
                    self.frames["automation"].set_progress(False)
                    self._log("All steps completed.")
                    self.running = False
                    self.frames["automation"].set_run_button_state(enabled=True, label="Run")
                elif typ == "failed_all":
                    self.frames["automation"].set_progress(False)
                    self._log("Pipeline stopped due to failure.")
                    self.running = False
                    self.frames["automation"].set_run_button_state(enabled=True, label="Run")
                elif typ == "reports_done":
                    self._log("Report job(s) completed.")
                    self.frames["reports"].set_reports_button_state(enabled=True, label="Create report(s)")
                elif typ == "reports_failed":
                    self._log("Report job(s) stopped due to failure.")
                    self.frames["reports"].set_reports_button_state(enabled=True, label="Create report(s)")
        except queue.Empty:
            pass
        finally:
            self.after(100, self._drain_queue)

# ============================== Screens ==============================
class MenuScreen(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        title = ttk.Label(self, text="Welcome", style="Title.TLabel")
        title.pack(anchor="center", pady=(10, 20))

        btns = ttk.Frame(self)
        btns.pack(expand=True)
        b1 = ttk.Button(btns, text="Data Automation", style="Big.TButton",
                        command=lambda: app.show("automation"))
        b2 = ttk.Button(btns, text="Reports", style="Big.TButton",
                        command=lambda: app.show("reports"))
        b1.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")
        b2.grid(row=0, column=1, padx=12, pady=12, sticky="nsew")
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        ttk.Label(self, text="Tip: Press Esc to return to this menu", foreground="#666").pack(anchor="center", pady=(8, 0))

class AutomationScreen(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app

        header = ttk.Frame(self); header.pack(fill="x")
        ttk.Button(header, text="← Back to Menu", style="Back.TButton",
                   command=lambda: app.show("menu")).pack(side="left")
        ttk.Label(header, text="Data Automation", style="Title.TLabel").pack(side="left", padx=12)
        self.btn_run = ttk.Button(header, text="Run", command=self.start_pipeline); self.btn_run.pack(side="right")

        # --- Treat credentials UI ---
        creds = ttk.Labelframe(self, text="Treat credentials")
        creds.pack(fill="x", pady=(8, 0))

        self.var_treat_user = tk.StringVar(value=os.getenv("TREAT_USERNAME", self.app.settings.get("TREAT_USERNAME", "yxu")))
        self.var_treat_pass = tk.StringVar(value=os.getenv("TREAT_PASSWORD", self.app.settings.get("TREAT_PASSWORD", "")))

        ttk.Label(creds, text="Username").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        e_user = ttk.Entry(creds, textvariable=self.var_treat_user, width=28)
        e_user.grid(row=0, column=1, padx=8, pady=6, sticky="we")

        ttk.Label(creds, text="Password").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        self._pw_entry = ttk.Entry(creds, textvariable=self.var_treat_pass, width=28, show="*")
        self._pw_entry.grid(row=0, column=3, padx=8, pady=6, sticky="we")

        def _toggle_show():
            self._pw_entry.configure(show="" if show_var.get() else "*")
        show_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(creds, text="Show", variable=show_var, command=_toggle_show)\
            .grid(row=0, column=4, padx=8, pady=6, sticky="w")

        def _save_treat_creds():
            u = self.var_treat_user.get().strip()
            p = self.var_treat_pass.get()
            _apply_treat_env(u, p)  # update process env for upcoming imports
            s = self.app.settings.copy() if hasattr(self.app, "settings") else {}
            s["TREAT_USERNAME"] = u
            s["TREAT_PASSWORD"] = p  # NOTE: plain text for simplicity
            _save_settings(s)
            self.app.settings = s
            self.app._log("Saved Treat credentials (env updated).")

        ttk.Button(creds, text="Save", command=_save_treat_creds)\
            .grid(row=0, column=5, padx=8, pady=6, sticky="w")

        # Enter to save
        e_user.bind("<Return>", lambda _e: _save_treat_creds())
        self._pw_entry.bind("<Return>", lambda _e: _save_treat_creds())

        # grid stretch
        creds.columnconfigure(1, weight=1)
        creds.columnconfigure(3, weight=1)

        # -------- Flow Chart (responsive) --------
        self.flow = tk.Canvas(self, height=300, bg="#ffffff", highlightthickness=1, highlightbackground="#dddddd")
        self.flow.pack(fill="x", pady=(10, 8))
        self.flow.bind("<Configure>", self._on_flow_resize)

        self._flow_nodes = {}                       # node_key -> {"rect": id, "text": id, "bbox": (...)}
        self._flow_states = {k: "pending" for k in FLOW_LABELS}  # store state per NODE
        self._FLOW_FILL = {"pending": "#f5f5f5", "running": "#d9e6ff", "done": "#dcf7ea", "fail": "#ffe8cc"}
        self._FLOW_EDGE = {"pending": "#bfbfbf", "running": STATUS_COLORS["running"],
                           "done": STATUS_COLORS["done"], "fail": STATUS_COLORS["fail"]}

        # Blinking (animation)
        self._blinking = set()      # node_keys currently blinking
        self._blink_on = False
        self.after(450, self._blink_tick)

        self._draw_flow()

        # -------- Scrollable Steps List --------
        steps_box = ttk.Labelframe(self, text="Pipeline steps")
        steps_box.pack(fill="x", expand=False, pady=(8, 0))

        # limit visible height; content scrolls if longer
        self.steps_scroll = ScrollableFrame(steps_box, height=180)
        self.steps_scroll.pack(fill="x", expand=True)

        self.rows = {}
        for key, title_text, _ in STEPS:
            row = ttk.Frame(self.steps_scroll.inner)
            row.pack(fill="x", pady=4)
            lt = ttk.Label(row, text=title_text, width=60)
            lt.pack(side="left")
            ls = ttk.Label(row, text="Pending", foreground=STATUS_COLORS["pending"])
            ls.pack(side="left", padx=8)
            self.rows[key] = {"title": lt, "status": ls, "frame": row}  # store row frame for auto-scroll

        ttk.Separator(self).pack(fill="x", pady=8)
        controls = ttk.Frame(self); controls.pack(fill="x")
        ttk.Button(controls, text="Close", command=self.app.destroy).pack(side="right")

        self.progress = ttk.Progressbar(self, mode="indeterminate"); self.progress.pack(fill="x")
        self.progress_running = False

    # ---------- Responsive layout ----------
    def _on_flow_resize(self, _event): self._draw_flow()

    def _layout(self):
        w = max(self.flow.winfo_width(), 900)
        h = max(self.flow.winfo_height(), 240)

        # 5 equal columns (equal arrow lengths)
        margin_x = int(w * 0.04)
        cols = 5
        col_space = (w - 2 * margin_x) / cols
        cx = [int(margin_x + col_space * (i + 0.5)) for i in range(cols)]

        # Rows
        top_y = int(h * 0.25)
        mid_y = int(h * 0.50)
        bottom_y = int(h * 0.75)

        # Larger boxes
        box_w = int(min(col_space * 0.80, 380))
        box_h = int(max(h * 0.26, 64))

        return {"cx": cx, "top_y": top_y, "mid_y": mid_y, "bottom_y": bottom_y,
                "box_w": box_w, "box_h": box_h}

    # ---------- Flow drawing ----------
    def _draw_node(self, node_key, center_x, center_y, w, h):
        label = FLOW_LABELS[node_key]
        x0, y0, x1, y1 = center_x - w//2, center_y - h//2, center_x + w//2, center_y + h//2
        rect = self.flow.create_rectangle(x0, y0, x1, y1,
                                          fill=self._FLOW_FILL["pending"],
                                          outline=self._FLOW_EDGE["pending"], width=2)
        text = self.flow.create_text(center_x, center_y, text=label,
                                     font=("Segoe UI", 10), fill="#222", width=w - 24)
        self._flow_nodes[node_key] = {"rect": rect, "text": text, "bbox": (x0, y0, x1, y1)}
        return rect

    def _draw_arrow(self, from_node, to_node):
        fx0, fy0, fx1, fy1 = self._flow_nodes[from_node]["bbox"]
        tx0, ty0, tx1, ty1 = self._flow_nodes[to_node]["bbox"]
        start = (fx1, (fy0 + fy1) // 2)
        end   = (tx0, (ty0 + ty1) // 2)

        # If same column, draw vertical arrow
        if abs(start[0] - end[0]) < 6:
            x = (fx1 + tx0) // 2
            y0 = (fy0 + fy1) // 2
            y1 = (ty0 + ty1) // 2
            self.flow.create_line(x, y0 + 6, x, y1 - 6, arrow=tk.LAST, width=2, fill="#888")
            return

        # elbow route if too close horizontally
        if start[0] + 16 >= end[0]:
            midx = max(fx1, tx1) + 30
            self.flow.create_line(start[0] + 8, start[1], midx, start[1], width=2, fill="#888")
            self.flow.create_line(midx, start[1], midx, end[1], width=2, fill="#888")
            self.flow.create_line(midx, end[1], end[0] - 8, end[1], arrow=tk.LAST, width=2, fill="#888")
        else:
            self.flow.create_line(start[0] + 8, start[1], end[0] - 8, end[1], arrow=tk.LAST, width=2, fill="#888")

    def _draw_flow(self):
        self.flow.delete("all")
        self._flow_nodes.clear()

        L = self._layout()
        cx, top_y, mid_y, bottom_y = L["cx"], L["top_y"], L["mid_y"], L["bottom_y"]
        bw, bh = L["box_w"], L["box_h"]

        # Column 0: Clean (middle)
        self._draw_node("clean", cx[0], mid_y, bw, bh)

        # Column 1: Treat (top), AlayaCare (bottom - same row as Email)
        self._draw_node("treat_box", cx[1], top_y,    bw, bh)
        self._draw_node("alayacare", cx[1], bottom_y, bw, bh)

        # Column 2: Downloads (top), Email (bottom)
        self._draw_node("downloads", cx[2], top_y,    bw, bh)
        self._draw_node("email",     cx[2], bottom_y, bw, bh)

        # Column 3: Raw Data (middle)
        self._draw_node("raw",       cx[3], mid_y,    bw, bh)

        # Column 4: Clean Data (middle)
        self._draw_node("clean_data",cx[4], mid_y,    bw, bh)

        # Arrows
        self._draw_arrow("clean", "treat_box")
        self._draw_arrow("clean", "alayacare")
        self._draw_arrow("treat_box", "downloads")
        self._draw_arrow("alayacare", "email")
        self._draw_arrow("downloads", "raw")
        self._draw_arrow("email", "raw")
        self._draw_arrow("raw", "clean_data")

        # Re-apply saved node states after redraw and blinking visuals
        for node_key, state in self._flow_states.items():
            self._set_node_state(node_key, state)
            self._apply_blink_visual(node_key)

    # ----- Node coloring & blinking -----
    def _set_node_state(self, node_key: str, state: str):
        node = self._flow_nodes.get(node_key)
        if not node:
            return
        self._flow_states[node_key] = state
        self.flow.itemconfig(
            node["rect"],
            fill=self._FLOW_FILL.get(state, self._FLOW_FILL["pending"]),
            outline=self._FLOW_EDGE.get(state, self._FLOW_EDGE["pending"]),
            width=2,
        )

    def _apply_blink_visual(self, node_key: str):
        """If node is blinking and running, apply the alternating outline."""
        if node_key not in self._blinking:
            return
        if self._flow_states.get(node_key) != "running":
            return
        node = self._flow_nodes.get(node_key)
        if not node:
            return
        edge = "#1f6feb" if self._blink_on else "#ffba08"
        self.flow.itemconfig(node["rect"], outline=edge, width=(3 if self._blink_on else 2))

    def _start_blink(self, node_key: str):
        self._blinking.add(node_key)
        self._apply_blink_visual(node_key)

    def _stop_blink(self, node_key: str):
        if node_key in self._blinking:
            self._blinking.remove(node_key)
        # restore standard outline for current state
        self._set_node_state(node_key, self._flow_states.get(node_key, "pending"))

    def _blink_tick(self):
        self._blink_on = not self._blink_on
        for node_key in list(self._blinking):
            self._apply_blink_visual(node_key)
        self.after(450, self._blink_tick)

    # ---------- Auto-scroll to running step ----------
    def _auto_scroll_to_step(self, step_key: str):
        """
        Center the running step row inside the scrollable viewport.
        """
        try:
            row = self.rows[step_key]["frame"]
            canvas = self.steps_scroll.canvas
            inner = self.steps_scroll.inner

            # ensure geometry is up-to-date
            self.update_idletasks()

            row_y = row.winfo_y()                    # y relative to inner frame
            canvas_h = canvas.winfo_height()
            inner_h = inner.winfo_height()

            if inner_h <= canvas_h:
                return  # nothing to scroll

            # target top so the row is roughly centered
            target_top = max(0, row_y - canvas_h // 2)
            frac = target_top / float(inner_h - canvas_h)
            frac = max(0.0, min(1.0, frac))
            canvas.yview_moveto(frac)
        except Exception:
            pass

    # API used by App queue
    def set_progress(self, on: bool):
        if on and not self.progress_running:
            self.progress.start(12); self.progress_running = True
        elif not on and self.progress_running:
            self.progress.stop(); self.progress_running = False

    def set_status(self, step_key: str, state: str, text: str | None = None):
        # update status row
        if step_key in self.rows:
            lab = self.rows[step_key]["status"]
            lab.configure(foreground=STATUS_COLORS[state])
            lab.configure(text=(text if text is not None else state.capitalize()))

        # auto-scroll when a step starts running
        if state == "running":
            # defer to end of event loop tick to ensure updated sizes
            self.after(0, lambda sk=step_key: self._auto_scroll_to_step(sk))

        # update all mapped nodes for this step
        node_keys = STEP_NODE_MAP.get(step_key, [])
        for node_key in node_keys:
            if step_key == "alaya" and state == "running":
                # Special: AlayaCare & Email turn green immediately (no blinking)
                self._stop_blink(node_key)
                self._set_node_state(node_key, "done")
                continue

            if state == "running":
                self._set_node_state(node_key, "running")
                self._start_blink(node_key)
            else:
                self._stop_blink(node_key)
                self._set_node_state(node_key, state)

    def set_run_button_state(self, enabled=True, label="Run"):
        self.btn_run.configure(state=("normal" if enabled else "disabled"), text=label)

    def reset_status(self):
        # reset rows
        for key in self.rows:
            self.rows[key]["status"].configure(text="Pending", foreground=STATUS_COLORS["pending"])
        # reset all node colors and blinking
        for node_key in FLOW_LABELS:
            self._stop_blink(node_key)
            self._set_node_state(node_key, "pending")
        self.set_progress(False)

    def start_pipeline(self):
        if self.app.running:
            return
        self.reset_status()
        self.app.running = True
        self.set_run_button_state(enabled=False, label="Running…")
        threading.Thread(target=run_pipeline, args=(self.app,), daemon=True).start()

class ReportsScreen(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app

        header = ttk.Frame(self); header.pack(fill="x")
        ttk.Button(header, text="← Back to Menu", style="Back.TButton",
                   command=lambda: app.show("menu")).pack(side="left")
        ttk.Label(header, text="Reports", style="Title.TLabel").pack(side="left", padx=12)

        body = ttk.Frame(self); body.pack(fill="both", expand=True, pady=(8, 0))

        self.report_vars = {}
        grid = ttk.Frame(body); grid.pack(fill="x", padx=8, pady=8)

        for i, (key, label, _module) in enumerate(REPORTS):
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(grid, text=label, variable=var)
            cb.grid(row=i, column=0, sticky="w", padx=4, pady=6)
            self.report_vars[key] = var

        controls = ttk.Frame(self); controls.pack(fill="x", pady=(8, 0))
        self.btn_create = ttk.Button(controls, text="Create report(s)", command=self.start_reports)
        self.btn_create.pack(side="left")

    def set_reports_button_state(self, enabled=True, label="Create report(s)"):
        self.btn_create.configure(state=("normal" if enabled else "disabled"), text=label)

    def start_reports(self):
        selected = [(key, label, module)
                    for (key, label, module) in REPORTS
                    if self.report_vars.get(key, tk.BooleanVar()).get()]
        if not selected:
            self.app._log("No reports selected."); return
        self.set_reports_button_state(enabled=False, label="Running…")
        threading.Thread(target=run_reports, args=(self.app, selected), daemon=True).start()

# ============================== Workers ==============================
def run_pipeline(app: App, stop_on_error: bool = True):
    # Initialize COM for this worker thread
    com_inited = False
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inited = True
    except Exception as e:
        app.q.put(("log", {"msg": f"Warning: COM init failed (pipeline): {e}"}))

    try:
        for key, title_text, module in STEPS:
            count, plain = _step_parts(title_text)
            start_msg = f"Starting: {plain}"
            print(start_msg); app.q.put(("log", {"msg": start_msg}))
            app.q.put(("status", {"key": key, "state": "running", "text": "Running…"}))
            t0 = time.time()

            try:
                runner = _get_runner(module)
            except Exception:
                dt = time.time() - t0
                app.q.put(("status", {"key": key, "state": "fail", "text": f"Failed to import ({dt:.1f}s)"}))
                app.q.put(("log", {"msg": traceback.format_exc()}))
                if stop_on_error:
                    app.q.put(("failed_all", {})); return
                continue

            try:
                runner()
                dt = time.time() - t0
                app.q.put(("status", {"key": key, "state": "done", "text": f"Done ({dt:.1f}s)"}))
                finish_msg = (f"[{count}] Finished" if count else "Finished") + f" ({dt:.1f}s)"
                print(finish_msg); app.q.put(("log", {"msg": finish_msg}))
            except Exception:
                dt = time.time() - t0
                app.q.put(("status", {"key": key, "state": "fail", "text": f"Failed ({dt:.1f}s)"}))
                app.q.put(("log", {"msg": traceback.format_exc()}))
                if stop_on_error:
                    app.q.put(("failed_all", {})); return
                # else continue
    finally:
        if com_inited:
            try: pythoncom.CoUninitialize()
            except Exception: pass
        app.q.put(("done_all", {}))

def run_reports(app: App, selected_reports: list[tuple[str,str,str]]):
    com_inited = False
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        com_inited = True
    except Exception as e:
        app.q.put(("log", {"msg": f"Warning: COM init failed (reports): {e}"}))

    try:
        for key, label, module in selected_reports:
            app.q.put(("log", {"msg": f"Starting report: {label}"}))
            t0 = time.time()
            runner = _get_runner(module)
            runner()
            dt = time.time() - t0
            app.q.put(("log", {"msg": f"Finished report: {label} ({dt:.1f}s)"}))
        app.q.put(("reports_done", {}))
    except Exception:
        app.q.put(("log", {"msg": traceback.format_exc()}))
        app.q.put(("reports_failed", {}))
    finally:
        if com_inited:
            try: pythoncom.CoUninitialize()
            except Exception: pass

# ============================== Entry ==============================
def main():
    ensure_shortcuts()
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
