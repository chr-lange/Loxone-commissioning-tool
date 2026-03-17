#!/usr/bin/env python3
"""
Loxone I/O Commissioning Checklist — Desktop GUI  v2.2
Run:  py loxone_checklist_gui.py
"""

APP_VERSION = "2.2"
CONFIG_FILE = "loxone_tool_config.json"   # saved next to the script

import sys, json, threading, os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# ── Optional deps ─────────────────────────────────────────────────────────────
missing = []
try:
    import requests
    from requests.auth import HTTPBasicAuth
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    missing.append("requests")
try:
    from reportlab.lib.pagesizes import A4
except ImportError:
    missing.append("reportlab")

if missing:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing dependencies",
        f"Install:\n\n  pip install {' '.join(missing)}\n\nThen restart.")
    sys.exit(1)

try:
    from loxone_checklist import (
        fetch_structure, load_local_structure, parse_structure, generate_pdf,
        CONTROL_CMDS, DIMMER_TYPES, PRIMARY_STATE,
    )
except ImportError as e:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing file",
        f"loxone_checklist.py must be in the same folder.\n{e}")
    sys.exit(1)

# ── Palette ───────────────────────────────────────────────────────────────────
BG      = "#f0f2f5"
CARD    = "#ffffff"
DARK    = "#1a3a5c"
GREEN   = "#6EBD44"
GREEN2  = "#5aaa33"
MID     = "#2e6da4"
RED     = "#c0392b"
RED2    = "#a93226"
ORANGE  = "#e67e22"
OK_C    = "#1e8449"
MUTED   = "#666666"
BORDER  = "#dde1e7"

# Status colours for tree
S_COLOR = {"ok": "#1e8449", "nok": "#c0392b", "skip": "#e67e22", None: "#555555"}
S_ICON  = {"ok": "✓", "nok": "✗", "skip": "→", None: "○"}
S_TAG   = {"ok": "s_ok", "nok": "s_nok", "skip": "s_skip", None: "s_none"}

FONT      = ("Segoe UI", 10)
FONT_B    = ("Segoe UI", 10, "bold")
FONT_TITLE= ("Segoe UI", 13, "bold")
FONT_SM   = ("Segoe UI", 9)
FONT_MONO = ("Consolas", 9)
FONT_BIG  = ("Segoe UI", 12, "bold")


# ── Buttons ───────────────────────────────────────────────────────────────────
class Btn(tk.Button):
    def __init__(self, parent, color=GREEN, hover=GREEN2, fg="#fff", **kw):
        kw.setdefault("font", FONT_B); kw.setdefault("relief", "flat")
        kw.setdefault("cursor", "hand2"); kw.setdefault("padx", 12)
        kw.setdefault("pady", 6)
        super().__init__(parent, bg=color, fg=fg, activebackground=hover,
                         activeforeground=fg, **kw)
        self._c = color; self._h = hover
        self.bind("<Enter>", lambda e: self.config(bg=hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._c))

    def set_color(self, color, hover=None):
        self._c = color
        self._h = hover or color
        self.config(bg=color)

class SmBtn(tk.Button):
    def __init__(self, parent, **kw):
        kw.setdefault("font", FONT_SM); kw.setdefault("relief", "flat")
        kw.setdefault("cursor", "hand2"); kw.setdefault("padx", 8)
        kw.setdefault("pady", 4)
        super().__init__(parent, bg="#e0e0e0", fg=DARK,
                         activebackground="#c8c8c8", **kw)

def _lbl(parent, text, **kw):
    kw.setdefault("font", FONT); kw.setdefault("fg", "#333")
    kw.setdefault("bg", parent.cget("bg"))
    return tk.Label(parent, text=text, **kw)


# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Loxone I/O Commissioning Tool  v{APP_VERSION}")
        self.configure(bg=BG)
        self.minsize(980, 660)
        self.geometry("1080x730")

        # ── Window icon ───────────────────────────────────────────
        try:
            if getattr(sys, "frozen", False):
                _icon_base = sys._MEIPASS          # PyInstaller temp dir
            else:
                _icon_base = os.path.dirname(os.path.abspath(__file__))
            _ico = os.path.join(_icon_base, "loxone_icon.ico")
            if os.path.isfile(_ico):
                self.iconbitmap(_ico)
        except Exception:
            pass   # icon is cosmetic — never crash over it

        # state
        self._structure   = None
        self._inputs      = []
        self._outputs     = []
        self._others      = []
        self._ms_info     = {}
        self._conn        = {}
        self._busy        = False
        self._status_map  = {}   # uuid → "ok"|"nok"|"skip"|None
        self._notes       = {}   # uuid → str
        self._ignored     = set()  # uuids excluded from reports
        self._tree_items  = {}   # iid  → item dict
        self._uuid_to_iid = {}   # uuid → iid
        self._selected_item = None
        self._all_items   = []   # flat list of all items for session

        # company / branding (persisted to config file)
        self._company_name = tk.StringVar()
        self._company_sub  = tk.StringVar()

        self._build_ui()
        self._load_config()
        self._center()

    # ─────────────────────────────────────────────────────────────────────────
    #  UI BUILD
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"Loxone  I/O  Commissioning  Tool  v{APP_VERSION}",
                 font=FONT_TITLE, fg="#fff", bg=DARK).pack(side="left", padx=16)
        self._hdr_right = tk.Frame(hdr, bg=DARK)
        self._hdr_right.pack(side="right", padx=16)
        self._hdr_company = tk.Label(self._hdr_right, text="", font=FONT_B,
                                     fg="#cce0f5", bg=DARK)
        self._hdr_company.pack(anchor="e")
        self._hdr_proj = tk.Label(self._hdr_right, text="", font=FONT_SM,
                                  fg="#99bbdd", bg=DARK)
        self._hdr_proj.pack(anchor="e")

        # Notebook
        s = ttk.Style(); s.theme_use("clam")
        s.configure("TNotebook",      background=BG, borderwidth=0)
        s.configure("TNotebook.Tab",  font=FONT_B, padding=[14,6],
                    background="#dde1e7", foreground=DARK)
        s.map("TNotebook.Tab",
              background=[("selected", CARD)],
              foreground=[("selected", DARK)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        t1 = tk.Frame(nb, bg=BG)
        t2 = tk.Frame(nb, bg=BG)
        t3 = tk.Frame(nb, bg=BG)
        nb.add(t1, text="  Connect  ")
        nb.add(t2, text="  Test I/O  ")
        nb.add(t3, text="  Report  ")

        self._build_connect(t1)
        self._build_test(t2)
        self._build_report(t3)

        # Log
        lf = tk.Frame(self, bg=CARD, bd=1, relief="solid")
        lf.pack(fill="x", padx=8, pady=(0,5))
        lh = tk.Frame(lf, bg=CARD); lh.pack(fill="x", padx=8, pady=(3,0))
        tk.Label(lh, text="Log", font=FONT_B, fg=DARK, bg=CARD).pack(side="left")
        SmBtn(lh, text="Clear", command=self._clear_log).pack(side="right")
        self._log = scrolledtext.ScrolledText(
            lf, height=4, font=FONT_MONO, bg="#1e1e1e", fg="#d4d4d4",
            relief="flat", state="disabled", wrap="word")
        self._log.pack(fill="x", padx=4, pady=(2,4))

        # Status bar
        bar = tk.Frame(self, bg="#dde1e7", pady=4)
        bar.pack(fill="x")
        self._status = tk.StringVar(value="Ready — connect to a Miniserver to start.")
        tk.Label(bar, textvariable=self._status,
                 font=FONT_SM, fg=MUTED, bg="#dde1e7").pack(side="left", padx=12)

    # ── TAB 1: Connect ────────────────────────────────────────────────────────
    def _build_connect(self, parent):
        outer = tk.Frame(parent, bg=BG, padx=16, pady=14)
        outer.pack(fill="both", expand=True)

        left  = tk.Frame(outer, bg=CARD, bd=1, relief="solid", padx=16, pady=14)
        right = tk.Frame(outer, bg=CARD, bd=1, relief="solid", padx=16, pady=14)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        right.grid(row=0, column=1, sticky="nsew")
        outer.columnconfigure(0, weight=2); outer.columnconfigure(1, weight=3)

        _lbl(left, "Miniserver Connection", font=FONT_B, fg=DARK).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0,10))

        self._src = tk.StringVar(value="live")
        sf = tk.Frame(left, bg=CARD); sf.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0,8))
        for val, txt in [("live","Live Miniserver"),("file","Local JSON file")]:
            tk.Radiobutton(sf, text=txt, variable=self._src, value=val,
                           command=self._toggle_src,
                           font=FONT, bg=CARD, activebackground=CARD).pack(side="left", padx=(0,12))

        # Live fields
        self._lf = tk.Frame(left, bg=CARD)
        self._lf.grid(row=2, column=0, columnspan=3, sticky="ew")
        self._cv = []
        for i,(lbl,ph,sec) in enumerate([
            ("IP / Hostname:","e.g. 192.168.1.100",False),
            ("Username:","admin",False),
            ("Password:","",True),
        ]):
            _lbl(self._lf, lbl).grid(row=i, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=ph if not sec else "")
            e = tk.Entry(self._lf, textvariable=var, width=24, font=FONT,
                         relief="solid", bd=1, show="•" if sec else "",
                         fg="#aaa" if ph and not sec else "#000")
            e.grid(row=i, column=1, sticky="ew", padx=(6,0), pady=3)
            if ph and not sec:
                e.bind("<FocusIn>",  lambda ev,w=e,p=ph: self._phi(w,p))
                e.bind("<FocusOut>", lambda ev,w=e,p=ph: self._pho(w,p))
            self._cv.append(var)
        self._lf.columnconfigure(1, weight=1)

        self._chk_https = tk.BooleanVar()
        self._chk_token = tk.BooleanVar()
        for row,var,txt in [(3,self._chk_https,"Use HTTPS"),
                            (4,self._chk_token,"Token auth (Gen2 Miniserver)")]:
            tk.Checkbutton(self._lf, text=txt, variable=var,
                           font=FONT_SM, bg=CARD, activebackground=CARD).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=(3,0))

        # File frame
        self._ff = tk.Frame(left, bg=CARD)
        _lbl(self._ff,"JSON file:").grid(row=0,column=0,sticky="w")
        self._fp = tk.StringVar()
        tk.Entry(self._ff, textvariable=self._fp, width=20,
                 font=FONT, relief="solid", bd=1).grid(row=0,column=1,padx=(6,4),sticky="ew")
        SmBtn(self._ff, text="Browse…", command=self._browse_json).grid(row=0,column=2)
        self._ff.columnconfigure(1, weight=1)

        Btn(left, text="Connect & Fetch Structure", command=self._connect).grid(
            row=10, column=0, columnspan=3, sticky="ew", pady=(14,4))
        self._stat_lbl = _lbl(left, "", font=FONT_SM, fg=OK_C)
        self._stat_lbl.grid(row=11, column=0, columnspan=3, sticky="w")

        # Right: project info
        _lbl(right, "Project Info", font=FONT_B, fg=DARK).pack(anchor="w", pady=(0,8))
        self._info = tk.Text(right, height=6, font=FONT_SM,
                             bg="#f9f9f9", relief="solid", bd=1,
                             state="disabled", wrap="word")
        self._info.pack(fill="x")

        ttk.Separator(right).pack(fill="x", pady=10)

        # Company / branding fields
        _lbl(right, "Report Branding", font=FONT_B, fg=DARK).pack(anchor="w", pady=(0,6))
        bf = tk.Frame(right, bg=CARD); bf.pack(fill="x")
        bf.columnconfigure(1, weight=1)
        _lbl(bf, "Company name:").grid(row=0, column=0, sticky="w", pady=3)
        tk.Entry(bf, textvariable=self._company_name, font=FONT,
                 relief="solid", bd=1).grid(row=0, column=1, sticky="ew", padx=(6,0), pady=3)
        _lbl(bf, "Sub-title / dept:").grid(row=1, column=0, sticky="w", pady=3)
        tk.Entry(bf, textvariable=self._company_sub, font=FONT,
                 relief="solid", bd=1).grid(row=1, column=1, sticky="ew", padx=(6,0), pady=3)
        SmBtn(right, text="💾 Save branding", command=self._save_config).pack(
            anchor="w", pady=(6,0))
        _lbl(right, "Shown in PDF header and window title bar.",
             font=FONT_SM, fg=MUTED).pack(anchor="w", pady=(2,0))

        # Update header live as user types
        self._company_name.trace_add("write", lambda *_: self._update_hdr_company())
        self._company_sub .trace_add("write", lambda *_: self._update_hdr_company())

    # ── TAB 2: Test I/O ───────────────────────────────────────────────────────
    def _build_test(self, parent):
        outer = tk.Frame(parent, bg=BG, padx=10, pady=8)
        outer.pack(fill="both", expand=True)

        pane = tk.PanedWindow(outer, orient="horizontal", bg=BG,
                              sashwidth=6, sashrelief="flat")
        pane.pack(fill="both", expand=True)

        # ── Left: tree panel ──────────────────────────────────────────────────
        tp = tk.Frame(pane, bg=CARD, bd=1, relief="solid")
        pane.add(tp, minsize=290)

        # Title + session buttons
        th = tk.Frame(tp, bg=CARD, padx=8, pady=6)
        th.pack(fill="x")
        _lbl(th, "I/O by Room", font=FONT_B, fg=DARK, bg=CARD).pack(side="left")
        SmBtn(th, text="💾 Save", command=self._save_session).pack(side="right", padx=(2,0))
        SmBtn(th, text="📂 Load", command=self._load_session).pack(side="right", padx=(2,0))

        # Progress bar
        prog_f = tk.Frame(tp, bg=CARD, padx=8); prog_f.pack(fill="x", pady=(0,2))
        self._prog_lbl = _lbl(prog_f, "Not loaded", font=FONT_SM, fg=MUTED, bg=CARD)
        self._prog_lbl.pack(anchor="w")
        self._prog_bar = ttk.Progressbar(prog_f, length=220, mode="determinate")
        self._prog_bar.pack(fill="x", pady=(2,0))

        # Type toggle: Outputs / Inputs / All
        ttf = tk.Frame(tp, bg=CARD, padx=6); ttf.pack(fill="x", pady=(4,0))
        _lbl(ttf, "Show:", bg=CARD, font=FONT_SM).pack(side="left")
        self._tree_kind = tk.StringVar(value="outputs")
        for val, txt in [("outputs","Outputs"), ("inputs","Inputs"), ("all","All")]:
            tk.Radiobutton(ttf, text=txt, variable=self._tree_kind, value=val,
                           command=self._filter_tree, font=FONT_SM,
                           bg=CARD, activebackground=CARD).pack(side="left", padx=3)

        # Search filter
        ff = tk.Frame(tp, bg=CARD, padx=6); ff.pack(fill="x", pady=(2,0))
        _lbl(ff, "Filter:", bg=CARD).pack(side="left")
        self._tree_filter = tk.StringVar()
        self._tree_filter.trace_add("write", lambda *_: self._filter_tree())
        tk.Entry(ff, textvariable=self._tree_filter, font=FONT_SM,
                 relief="solid", bd=1).pack(side="left", fill="x", expand=True, padx=(4,0))

        # Status filter
        sf2 = tk.Frame(tp, bg=CARD, padx=6); sf2.pack(fill="x", pady=(2,4))
        _lbl(sf2, "Status:", bg=CARD, font=FONT_SM).pack(side="left")
        self._show_filter = tk.StringVar(value="all")
        for val, txt in [("all","All"),("none","Untested"),("ok","OK"),
                         ("nok","Not OK"),("ign","Ignored")]:
            tk.Radiobutton(sf2, text=txt, variable=self._show_filter, value=val,
                           command=self._filter_tree, font=FONT_SM,
                           bg=CARD, activebackground=CARD).pack(side="left", padx=2)

        # Tree
        ts = ttk.Style()
        ts.configure("Out.Treeview", font=FONT_SM, rowheight=22)
        ts.configure("Out.Treeview.Heading", font=FONT_B)
        ts.map("Out.Treeview", background=[("selected","#c8dff5")],
               foreground=[("selected", DARK)])

        sc = tk.Scrollbar(tp, orient="vertical")
        self._tree = ttk.Treeview(tp, style="Out.Treeview",
                                  yscrollcommand=sc.set,
                                  selectmode="browse", show="tree")
        sc.config(command=self._tree.yview)
        sc.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True, padx=2, pady=(0,4))
        self._tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # ── Right: control panel ──────────────────────────────────────────────
        cp = tk.Frame(pane, bg=CARD, bd=1, relief="solid")
        pane.add(cp, minsize=440)

        # Item header
        ih = tk.Frame(cp, bg=CARD, padx=12, pady=8); ih.pack(fill="x")
        self._ctrl_name = _lbl(ih, "— select an item from the list —",
                                font=FONT_B, fg=DARK, bg=CARD)
        self._ctrl_name.pack(anchor="w")
        self._ctrl_meta = _lbl(ih, "", font=FONT_SM, fg=MUTED, bg=CARD)
        self._ctrl_meta.pack(anchor="w")

        ttk.Separator(cp).pack(fill="x", padx=8)

        # ── STATUS row (primary action) ───────────────────────────────────────
        sr = tk.Frame(cp, bg="#f5f8fc", bd=1, relief="solid", padx=12, pady=10)
        sr.pack(fill="x", padx=8, pady=8)
        _lbl(sr, "Test Result:", font=FONT_B, fg=DARK, bg="#f5f8fc").grid(
            row=0, column=0, sticky="w", pady=(0,6), columnspan=5)

        self._btn_ok   = Btn(sr, text="✓  OK",      color=OK_C,   hover="#166a38",
                             font=FONT_BIG, padx=20, pady=10,
                             command=lambda: self._set_status_item("ok"))
        self._btn_nok  = Btn(sr, text="✗  Not OK",  color=RED,    hover=RED2,
                             font=FONT_BIG, padx=20, pady=10,
                             command=lambda: self._set_status_item("nok"))
        self._btn_skip = Btn(sr, text="→  Skip",    color=ORANGE, hover="#ca6f1e",
                             font=FONT_B, padx=14, pady=10,
                             command=lambda: self._set_status_item("skip"))
        self._btn_ign  = SmBtn(sr, text="⊘ Ignore",
                               command=self._toggle_ignore)
        self._btn_clr  = SmBtn(sr, text="Clear",
                               command=lambda: self._set_status_item(None))

        self._btn_ok.grid(  row=1, column=0, padx=(0,4))
        self._btn_nok.grid( row=1, column=1, padx=(0,4))
        self._btn_skip.grid(row=1, column=2, padx=(0,8))
        self._btn_ign.grid( row=1, column=3, padx=(0,4))
        self._btn_clr.grid( row=1, column=4)

        # Current status display
        self._curr_status = _lbl(sr, "", font=FONT_B, fg=MUTED, bg="#f5f8fc")
        self._curr_status.grid(row=2, column=0, columnspan=5, sticky="w", pady=(6,0))

        # Note field
        nf = tk.Frame(sr, bg="#f5f8fc")
        nf.grid(row=3, column=0, columnspan=5, sticky="ew", pady=(6,0))
        _lbl(nf, "Note / fault description:", font=FONT_SM, fg=MUTED, bg="#f5f8fc").pack(anchor="w")
        self._note_var = tk.StringVar()
        self._note_entry = tk.Entry(nf, textvariable=self._note_var, font=FONT,
                                    relief="solid", bd=1)
        self._note_entry.pack(fill="x", pady=(2,0))
        self._note_entry.bind("<Return>",   lambda _: self._save_note())
        self._note_entry.bind("<FocusOut>", lambda _: self._save_note())

        ttk.Separator(cp).pack(fill="x", padx=8, pady=(0,4))

        # ── Live value panel ──────────────────────────────────────────────────
        self._val_frame = tk.Frame(cp, bg="#f0f4f8", bd=1, relief="solid", padx=12, pady=6)
        vf = self._val_frame
        vf.pack(fill="x", padx=8, pady=(0,6))
        vf_hdr = tk.Frame(vf, bg="#f0f4f8"); vf_hdr.pack(fill="x")
        _lbl(vf_hdr, "Current Value:", font=FONT_SM, fg=MUTED, bg="#f0f4f8").pack(side="left")
        self._btn_read = SmBtn(vf_hdr, text="📡 Read",
                               command=self._read_value)
        self._btn_read.pack(side="right")
        self._val_display = _lbl(vf, "—", font=("Segoe UI", 14, "bold"),
                                 fg=DARK, bg="#f0f4f8")
        self._val_display.pack(anchor="w", pady=(4,0))
        self._val_all = _lbl(vf, "", font=FONT_SM, fg=MUTED, bg="#f0f4f8")
        self._val_all.pack(anchor="w")

        ttk.Separator(cp).pack(fill="x", padx=8, pady=(4,4))

        # ── Control commands ──────────────────────────────────────────────────
        self._btn_frame = tk.Frame(cp, bg=CARD, padx=12, pady=4)
        self._btn_frame.pack(fill="x")
        _lbl(self._btn_frame, "Select an item to see controls.",
             font=FONT_SM, fg=MUTED, bg=CARD).pack(anchor="w")

        # Dimmer slider
        self._dim_frame = tk.Frame(cp, bg=CARD, padx=12, pady=4)
        self._dim_label = _lbl(self._dim_frame, "Set value (0–100):", bg=CARD)
        self._dim_label.pack(anchor="w")
        dr = tk.Frame(self._dim_frame, bg=CARD); dr.pack(fill="x", pady=4)
        self._dim_var = tk.IntVar(value=50)
        tk.Scale(dr, from_=0, to=100, orient="horizontal",
                 variable=self._dim_var, font=FONT_SM, bg=CARD,
                 highlightthickness=0, troughcolor=BORDER,
                 activebackground=MID).pack(side="left", fill="x", expand=True)
        Btn(dr, text="Set", color=MID, hover="#255e8e", pady=4,
            command=self._send_dim).pack(side="left", padx=(8,0))

        # Command log
        ttk.Separator(cp).pack(fill="x", padx=8, pady=(8,4))
        _lbl(cp, "Command log:", font=FONT_SM, fg=MUTED, bg=CARD).pack(
            anchor="w", padx=12)
        self._cmd_log = scrolledtext.ScrolledText(
            cp, height=6, font=FONT_MONO, bg="#1e1e1e", fg="#d4d4d4",
            relief="flat", state="disabled", wrap="word")
        self._cmd_log.pack(fill="both", expand=True, padx=6, pady=(0,6))

    # ── TAB 3: Report ─────────────────────────────────────────────────────────
    def _build_report(self, parent):
        outer = tk.Frame(parent, bg=BG, padx=16, pady=14)
        outer.pack(fill="both", expand=True)

        left  = tk.Frame(outer, bg=CARD, bd=1, relief="solid", padx=14, pady=12)
        right = tk.Frame(outer, bg=CARD, bd=1, relief="solid", padx=14, pady=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        right.grid(row=0, column=1, sticky="nsew")
        outer.columnconfigure(0, weight=2); outer.columnconfigure(1, weight=3)

        # ── Left: filters ─────────────────────────────────────────────────────
        _lbl(left, "Report Filters", font=FONT_B, fg=DARK).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0,8))

        _lbl(left, "Sections:", font=FONT_SM, fg=MUTED).grid(
            row=1, column=0, columnspan=2, sticky="w")
        self._inc_inputs  = tk.BooleanVar(value=True)
        self._inc_outputs = tk.BooleanVar(value=True)
        self._inc_other   = tk.BooleanVar(value=False)
        for row,var,txt in [
            (2, self._inc_inputs,  "✔  Inputs  (sensors / digital / analog)"),
            (3, self._inc_outputs, "✔  Outputs  (actors / lights / blinds)"),
            (4, self._inc_other,   "✔  Other control types"),
        ]:
            tk.Checkbutton(left, text=txt, variable=var, font=FONT_SM,
                           bg=CARD, activebackground=CARD).grid(
                row=row, column=0, columnspan=2, sticky="w", padx=(10,0))

        ttk.Separator(left).grid(row=5, column=0, columnspan=2, sticky="ew", pady=8)

        _lbl(left, "Status filter:", font=FONT_SM, fg=MUTED).grid(
            row=6, column=0, columnspan=2, sticky="w")
        self._rpt_status = tk.StringVar(value="all")
        for row,val,txt in [
            (7, "all",  "All items"),
            (8, "ok",   "OK items only"),
            (9, "nok",  "Not OK items only"),
            (10,"none", "Untested items only"),
        ]:
            tk.Radiobutton(left, text=txt, variable=self._rpt_status, value=val,
                           command=self._update_preview,
                           font=FONT_SM, bg=CARD, activebackground=CARD).grid(
                row=row, column=0, columnspan=2, sticky="w", padx=(10,0))

        ttk.Separator(left).grid(row=11, column=0, columnspan=2, sticky="ew", pady=8)

        _lbl(left, "Rooms (select to include):", font=FONT_SM, fg=MUTED).grid(
            row=12, column=0, columnspan=2, sticky="w")

        rf = tk.Frame(left, bg=CARD)
        rf.grid(row=13, column=0, columnspan=2, sticky="ew", pady=(4,0))
        rsc = tk.Scrollbar(rf); rsc.pack(side="right", fill="y")
        self._room_list = tk.Listbox(rf, selectmode="multiple", font=FONT_SM, height=7,
                                     yscrollcommand=rsc.set, relief="solid", bd=1,
                                     selectbackground=MID, selectforeground="#fff",
                                     exportselection=False)
        self._room_list.pack(fill="both", expand=True)
        rsc.config(command=self._room_list.yview)

        br = tk.Frame(left, bg=CARD)
        br.grid(row=14, column=0, columnspan=2, sticky="w", pady=(4,0))
        SmBtn(br, text="All",  command=self._rooms_all).pack(side="left")
        SmBtn(br, text="None", command=self._rooms_none).pack(side="left", padx=4)

        # ── Right: output options ─────────────────────────────────────────────
        _lbl(right, "PDF Output", font=FONT_B, fg=DARK).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0,8))

        _lbl(right, "Save as:").grid(row=1, column=0, sticky="w", pady=3)
        self._out_path = tk.StringVar(value=str(Path.home()/"loxone_checklist.pdf"))
        tk.Entry(right, textvariable=self._out_path, width=24,
                 font=FONT, relief="solid", bd=1).grid(row=1, column=1, sticky="ew", padx=(6,4))
        SmBtn(right, text="Browse…", command=self._browse_out).grid(row=1, column=2)

        ttk.Separator(right).grid(row=2, column=0, columnspan=3, sticky="ew", pady=10)

        _lbl(right, "Will include:", font=FONT_SM, fg=MUTED).grid(
            row=3, column=0, columnspan=3, sticky="w")
        self._preview_lbl = _lbl(right, "—", font=FONT_SM, fg=DARK)
        self._preview_lbl.grid(row=4, column=0, columnspan=3, sticky="w", padx=(10,0), pady=(2,10))

        ttk.Separator(right).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0,10))

        self._gen_full = Btn(right, text="⬇  Generate Full Report",
                             command=lambda: self._generate("full"), state="disabled")
        self._gen_full.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0,6))

        self._gen_sel = Btn(right, text="⬇  Generate Filtered Report",
                            command=lambda: self._generate("filtered"),
                            color=MID, hover="#255e8e", state="disabled")
        self._gen_sel.grid(row=7, column=0, columnspan=3, sticky="ew")

        right.columnconfigure(1, weight=1)
        for v in [self._inc_inputs, self._inc_outputs, self._inc_other]:
            v.trace_add("write", lambda *_: self._update_preview())
        self._room_list.bind("<<ListboxSelect>>", lambda _: self._update_preview())
        self._rpt_status.trace_add("write", lambda *_: self._update_preview())

    # ─────────────────────────────────────────────────────────────────────────
    #  CONNECTION
    # ─────────────────────────────────────────────────────────────────────────
    def _toggle_src(self):
        if self._src.get() == "live":
            self._ff.grid_remove()
            self._lf.grid(row=2, column=0, columnspan=3, sticky="ew")
        else:
            self._lf.grid_remove()
            self._ff.grid(row=2, column=0, columnspan=3, sticky="ew")

    def _phi(self,w,ph):
        if w.get()==ph: w.delete(0,"end"); w.config(fg="#000")
    def _pho(self,w,ph):
        if not w.get(): w.insert(0,ph); w.config(fg="#aaa")

    def _browse_json(self):
        p = filedialog.askopenfilename(title="Select LoxAPP3.json",
            filetypes=[("JSON","*.json"),("All","*.*")])
        if p: self._fp.set(p)

    def _browse_out(self):
        p = filedialog.asksaveasfilename(title="Save PDF as…",
            defaultextension=".pdf", filetypes=[("PDF","*.pdf"),("All","*.*")])
        if p: self._out_path.set(p)

    def _connect(self):
        if self._busy: return
        src = self._src.get()
        if src == "file":
            p = self._fp.get().strip()
            if not p:
                messagebox.showwarning("No file","Browse to a LoxAPP3.json file first.")
                return
            self._conn = {}
            self._run_thread(self._do_load_file, p)
        else:
            host = self._cv[0].get().strip()
            user = self._cv[1].get().strip()
            pwd  = self._cv[2].get().strip()
            if host in ("","e.g. 192.168.1.100"):
                messagebox.showwarning("Missing","Enter Miniserver IP / hostname."); return
            self._conn = dict(host=host, username=user, password=pwd,
                              use_https=self._chk_https.get())
            self._run_thread(self._do_connect, host, user, pwd)

    def _do_load_file(self, path):
        self._set_status("Loading file…")
        self._log_msg(f"Loading: {path}", "info")
        try:
            s = load_local_structure(path)
            self.after(0, self._on_loaded, s)
        except SystemExit:
            self.after(0, self._on_err, "Cannot load file.")

    def _do_connect(self, host, user, pwd):
        self._set_status(f"Connecting to {host}…")
        self._log_msg(f"Connecting → {host}  ({user})", "info")
        try:
            s = fetch_structure(host, user, pwd,
                                use_https=self._chk_https.get(),
                                use_token=self._chk_token.get())
            self.after(0, self._on_loaded, s)
        except SystemExit:
            self.after(0, self._on_err, "Connection failed — check IP and credentials.")

    def _on_loaded(self, structure):
        self._structure = structure
        self._ms_info   = structure.get("msInfo", {})
        self._inputs, self._outputs, self._others = parse_structure(structure)
        # Tag each item with its kind for display purposes
        for item in self._inputs:  item["_kind"] = "input"
        for item in self._outputs: item["_kind"] = "output"
        for item in self._others:  item["_kind"] = "other"
        self._all_items = self._inputs + self._outputs + self._others
        # Preserve any existing status from a loaded session
        for item in self._all_items:
            if item["uuid"] not in self._status_map:
                self._status_map[item["uuid"]] = None
            if item["uuid"] not in self._notes:
                self._notes[item["uuid"]] = ""

        mi = self._ms_info
        self._set_info("\n".join(filter(None,[
            f"Project:    {mi.get('projectName','')}",
            f"Miniserver: {mi.get('msName','')}",
            f"Serial:     {mi.get('serialNr','')}",
            f"Firmware:   {mi.get('swVersion','')}",
            f"Location:   {mi.get('location','')}",
        ])))
        self._hdr_proj.config(text=mi.get("projectName",""))
        self._stat_lbl.config(
            text=f"  {len(self._inputs)} inputs   {len(self._outputs)} outputs   "
                 f"{len(self._others)} other", fg=OK_C)
        self._log_msg(f"Loaded OK — {len(self._all_items)} controls total", "success")
        self._set_status("Connected. Use Test I/O tab to test and mark items.")

        self._populate_tree()
        self._populate_room_list()
        self._gen_full.config(state="normal")
        self._gen_sel.config(state="normal")
        self._update_progress()
        self._update_preview()
        self._busy = False

    def _on_err(self, msg):
        self._log_msg(f"ERROR: {msg}", "error")
        self._set_status("Error — see log.")
        messagebox.showerror("Error", msg)
        self._busy = False

    # ─────────────────────────────────────────────────────────────────────────
    #  TREE  (shows inputs, outputs, or both)
    # ─────────────────────────────────────────────────────────────────────────
    def _visible_items(self):
        """Return items list based on the kind toggle."""
        kind = self._tree_kind.get()
        if kind == "outputs": return self._outputs
        if kind == "inputs":  return self._inputs
        return self._outputs + self._inputs   # "all"

    def _item_tag(self, uuid):
        if uuid in self._ignored:  return "s_ign"
        return S_TAG[self._status_map.get(uuid)]

    def _populate_tree(self):
        self._tree.delete(*self._tree.get_children())
        self._tree_items  = {}
        self._uuid_to_iid = {}
        by_room = defaultdict(list)
        for item in self._visible_items():
            by_room[item["room"]].append(item)
        for room in sorted(by_room):
            riid = self._tree.insert("", "end", text=f"  📁 {room}",
                                      open=True, tags=("room",))
            for item in sorted(by_room[room], key=lambda x: x["name"]):
                iid = self._tree.insert(riid, "end",
                    text=self._item_label(item),
                    tags=(self._item_tag(item["uuid"]),))
                self._tree_items[iid]           = item
                self._uuid_to_iid[item["uuid"]] = iid
        self._apply_tree_tags()

    def _item_label(self, item):
        uuid = item["uuid"]
        if uuid in self._ignored:
            icon = "⊘"
        else:
            st   = self._status_map.get(uuid)
            icon = S_ICON[st]
        kind_mark = "←" if item.get("_kind") == "input" else "→"
        bus  = f" [{item['bus']}]" if item.get("bus") else ""
        note = f"  ← {self._notes[uuid]}" if self._notes.get(uuid) else ""
        return f"  {icon} {kind_mark} {item['name']}  —  {item['type_label']}{bus}{note}"

    def _apply_tree_tags(self):
        self._tree.tag_configure("room",   foreground=DARK, font=FONT_B)
        self._tree.tag_configure("s_ok",   foreground=OK_C)
        self._tree.tag_configure("s_nok",  foreground=RED)
        self._tree.tag_configure("s_skip", foreground=ORANGE)
        self._tree.tag_configure("s_none", foreground="#555")
        self._tree.tag_configure("s_ign",  foreground="#aaaaaa")

    def _filter_tree(self):
        q  = self._tree_filter.get().lower().strip()
        sf = self._show_filter.get()
        self._tree.delete(*self._tree.get_children())
        self._tree_items  = {}
        self._uuid_to_iid = {}
        by_room = defaultdict(list)
        for item in self._visible_items():
            uuid = item["uuid"]
            ign  = uuid in self._ignored
            st   = self._status_map.get(uuid)
            # status filter
            if sf == "ign"  and not ign:          continue
            if sf != "all" and sf != "ign":
                if ign:                           continue   # hide ignored in non-ign filters
                if sf == "none" and st is not None: continue
                elif sf == "ok"  and st != "ok":    continue
                elif sf == "nok" and st != "nok":   continue
            # text filter
            if q and not any(q in str(v).lower() for v in
                             [item["name"], item["type_label"],
                              item["room"], item.get("bus","")]):
                continue
            by_room[item["room"]].append(item)
        for room in sorted(by_room):
            riid = self._tree.insert("", "end", text=f"  📁 {room}",
                                      open=True, tags=("room",))
            for item in sorted(by_room[room], key=lambda x: x["name"]):
                iid = self._tree.insert(riid, "end",
                    text=self._item_label(item),
                    tags=(self._item_tag(item["uuid"]),))
                self._tree_items[iid]           = item
                self._uuid_to_iid[item["uuid"]] = iid
        self._apply_tree_tags()

    def _refresh_tree_item(self, uuid):
        iid = self._uuid_to_iid.get(uuid)
        if iid and self._tree.exists(iid):
            item = self._tree_items[iid]
            self._tree.item(iid,
                text=self._item_label(item),
                tags=(self._item_tag(uuid),))
            self._apply_tree_tags()

    # ─────────────────────────────────────────────────────────────────────────
    #  ITEM SELECTION & STATUS
    # ─────────────────────────────────────────────────────────────────────────
    def _on_tree_select(self, _=None):
        sel = self._tree.selection()
        if not sel: return
        item = self._tree_items.get(sel[0])
        if not item: return
        self._selected_item = item
        uuid = item["uuid"]
        bus  = f" [{item['bus']}]" if item.get("bus") else ""
        kind = item.get("_kind", "output")
        kind_str = "INPUT  ←" if kind == "input" else "OUTPUT  →"
        self._ctrl_name.config(text=item["name"])
        self._ctrl_meta.config(
            text=f"{kind_str}   {item['type_label']}{bus}  ·  {item['room']}  ·  {item['category']}")

        self._note_var.set(self._notes.get(uuid, ""))
        self._refresh_status_display(uuid)

        # Live value panel — inputs only
        if kind == "input":
            self._val_frame.pack(fill="x", padx=8, pady=(0,6))
            self._val_display.config(text="—", fg=DARK)
            self._val_all.config(text="")
            if self._conn:
                self._run_thread(self._do_read_value, item)
        else:
            self._val_frame.pack_forget()

        # ── Command buttons ────────────────────────────────────────────────
        for w in self._btn_frame.winfo_children(): w.destroy()
        cmds = CONTROL_CMDS.get(item["type"], [])
        if cmds:
            _lbl(self._btn_frame, "Send command to Miniserver:",
                 font=FONT_SM, fg=MUTED, bg=CARD).pack(anchor="w", pady=(0,4))
            rf = tk.Frame(self._btn_frame, bg=CARD); rf.pack(fill="x")
            for i, (lbl, cmd) in enumerate(cmds):
                if cmd in ("off","stop","disarm","pause","3","4"):
                    c, h = "#7f8c8d","#6b7a7b"
                elif cmd in ("on","open","arm","play","1","0"):
                    c, h = OK_C,"#166a38"
                else:
                    c, h = MID,"#255e8e"
                Btn(rf, text=lbl, color=c, hover=h,
                    command=lambda c=cmd: self._send_cmd(c),
                    pady=5, padx=10).grid(row=0, column=i, padx=2, pady=2)
        elif kind == "input":
            _lbl(self._btn_frame,
                 "Input sensor — read only. Verify physically on site.",
                 font=FONT_SM, fg=MUTED, bg=CARD).pack(anchor="w")
        else:
            _lbl(self._btn_frame, "No commands defined for this type.",
                 font=FONT_SM, fg=MUTED, bg=CARD).pack(anchor="w")

        if item["type"] in DIMMER_TYPES:
            self._dim_frame.pack(fill="x", padx=12, pady=4)
        else:
            self._dim_frame.pack_forget()

    def _set_status_item(self, status):
        item = self._selected_item
        if not item: return
        # Remove from ignored if explicitly setting a status
        if status is not None:
            self._ignored.discard(item["uuid"])
        self._status_map[item["uuid"]] = status
        self._refresh_status_display(item["uuid"])
        self._refresh_tree_item(item["uuid"])
        self._update_progress()
        self._advance_to_next()

    def _toggle_ignore(self):
        item = self._selected_item
        if not item: return
        uuid = item["uuid"]
        if uuid in self._ignored:
            self._ignored.discard(uuid)
        else:
            self._ignored.add(uuid)
            self._status_map[uuid] = None   # clear status when ignoring
        self._refresh_status_display(uuid)
        self._refresh_tree_item(uuid)
        self._update_progress()
        self._advance_to_next()

    def _advance_to_next(self):
        sel = self._tree.selection()
        if not sel: return
        current_iid = sel[0]

        def _collect_leaves(parent=""):
            leaves = []
            for iid in self._tree.get_children(parent):
                if self._tree.get_children(iid):
                    leaves.extend(_collect_leaves(iid))
                else:
                    leaves.append(iid)
            return leaves

        leaves = _collect_leaves()
        if current_iid not in leaves: return
        idx = leaves.index(current_iid)
        if idx + 1 < len(leaves):
            next_iid = leaves[idx + 1]
            self._tree.selection_set(next_iid)
            self._tree.see(next_iid)
            self._on_tree_select()

    def _refresh_status_display(self, uuid):
        if uuid in self._ignored:
            self._curr_status.config(text="⊘  Ignored — excluded from report", fg="#aaaaaa")
            return
        st = self._status_map.get(uuid)
        if st is None:
            self._curr_status.config(text="Not tested", fg=MUTED)
        elif st == "ok":
            self._curr_status.config(text="✓  OK", fg=OK_C)
        elif st == "nok":
            self._curr_status.config(text="✗  Not OK", fg=RED)
        elif st == "skip":
            self._curr_status.config(text="→  Skipped", fg=ORANGE)

    def _save_note(self):
        item = self._selected_item
        if not item: return
        self._notes[item["uuid"]] = self._note_var.get().strip()
        self._refresh_tree_item(item["uuid"])

    # ─────────────────────────────────────────────────────────────────────────
    #  SEND COMMANDS
    # ─────────────────────────────────────────────────────────────────────────
    def _send_cmd(self, cmd):
        item = self._selected_item
        if not item: return
        if not self._conn:
            self._cmd_log_msg("No live connection — cannot send commands to Miniserver.", "error")
            return
        self._run_thread(self._do_send, item, cmd)

    def _send_dim(self):
        self._send_cmd(str(self._dim_var.get()))

    def _do_send(self, item, cmd):
        h     = self._conn
        proto = "https" if h.get("use_https") else "http"
        url   = f"{proto}://{h['host']}/jdev/sps/io/{item['uuid']}/{cmd}"
        self.after(0, self._cmd_log_msg, f"→ {item['name']}  [{cmd}]  …", "info")
        try:
            r = requests.get(url, auth=HTTPBasicAuth(h["username"], h["password"]),
                             timeout=8, verify=False)
            r.raise_for_status()
            resp = r.json()
            code = str(resp.get("LL",{}).get("Code","200"))
            if code in ("200","0"):
                self.after(0, self._cmd_log_msg,
                           f"✓ {item['name']}  →  [{cmd}]  (ok)", "success")
            else:
                self.after(0, self._cmd_log_msg,
                           f"✗ Miniserver code {code}  →  {item['name']}  [{cmd}]", "error")
        except Exception as e:
            self.after(0, self._cmd_log_msg, f"✗ {e}", "error")
        finally:
            self._busy = False

    # ─────────────────────────────────────────────────────────────────────────
    #  LIVE VALUE READ
    # ─────────────────────────────────────────────────────────────────────────
    def _read_value(self):
        """Triggered by the 📡 Read button."""
        item = self._selected_item
        if not item:
            return
        if not self._conn:
            self.after(0, lambda: self._val_display.config(text="No connection", fg=RED))
            return
        self._run_thread(self._do_read_value, item)

    def _do_read_value(self, item):
        """Background thread: poll live value for input, trying multiple endpoints.

        Loxone HTTP API strategies tried in order:
          1. GET /jdev/sps/getvalue/<normalized-state-uuid>
          2. GET /jdev/sps/getvalue/<control-uuid>
          3. GET /jdev/sps/io/<control-uuid>/state
        """
        h         = self._conn
        proto     = "https" if h.get("use_https") else "http"
        base      = f"{proto}://{h['host']}"
        auth      = HTTPBasicAuth(h["username"], h["password"])
        states    = item.get("states", {})
        ctrl_uuid = item["uuid"]
        primary_key = PRIMARY_STATE.get(item["type"])

        def _fetch(url):
            """Returns (value_str, http_status) — value_str is None on any failure."""
            try:
                r = requests.get(url, auth=auth, timeout=6, verify=False)
                if r.status_code not in (200, 201):
                    return None, r.status_code
                data = r.json()
                lox_code = str(data.get("LL", {}).get("Code", "200"))
                if lox_code not in ("200", "0"):
                    return None, f"lox{lox_code}"
                val = data.get("LL", {}).get("value")
                return (str(val).strip() if val is not None else None), r.status_code
            except Exception as exc:
                return None, str(exc)

        def _fmt(raw, state_name=""):
            try:
                f = float(raw)
                if f in (0.0, 1.0) and state_name.lower() in (
                        "active", "locked", "running", "open", "alarm",
                        "armed", "engaged", "onoff", "output"):
                    return "ON" if f == 1.0 else "OFF"
                return str(int(f)) if f == int(f) else f"{f:.2f}"
            except (ValueError, OverflowError):
                return raw

        results = {}   # state_name → formatted display value

        # ── Strategy 1: per-state UUID via getvalue ───────────────────────────
        for name, suuid in states.items():
            api_uuid = App._norm_uuid(suuid)
            val, _ = _fetch(f"{base}/jdev/sps/getvalue/{api_uuid}")
            if val is not None:
                results[name] = _fmt(val, name)

        # ── Strategy 2: control UUID via getvalue ─────────────────────────────
        if not results:
            val, status2 = _fetch(f"{base}/jdev/sps/getvalue/{ctrl_uuid}")
            if val is not None:
                results[primary_key or "value"] = _fmt(val, primary_key or "")

        # ── Strategy 3: io/<uuid>/state ───────────────────────────────────────
        if not results:
            val, status3 = _fetch(f"{base}/jdev/sps/io/{ctrl_uuid}/state")
            if val is not None:
                results[primary_key or "state"] = _fmt(val, primary_key or "")

        # ── Diagnostics if still nothing ──────────────────────────────────────
        if not results and states:
            first_uuid = App._norm_uuid(next(iter(states.values())))
            _, s1 = _fetch(f"{base}/jdev/sps/getvalue/{first_uuid}")
            _, s2 = _fetch(f"{base}/jdev/sps/getvalue/{ctrl_uuid}")
            _, s3 = _fetch(f"{base}/jdev/sps/io/{ctrl_uuid}/state")
            self.after(0, self._log_msg,
                f"⚠ {item['name']}: all read attempts failed — "
                f"state-uuid→HTTP {s1}  ctrl-uuid→HTTP {s2}  io/state→HTTP {s3}", "error")

        primary_raw = (results.get(primary_key)
                       or (next(iter(results.values())) if results else "N/A"))
        all_text = "  |  ".join(f"{k}: {v}" for k, v in results.items())

        self.after(0, lambda: self._val_display.config(text=primary_raw, fg=DARK))
        self.after(0, lambda: self._val_all.config(text=all_text))
        self._busy = False

    # ─────────────────────────────────────────────────────────────────────────
    #  PROGRESS
    # ─────────────────────────────────────────────────────────────────────────
    def _update_progress(self):
        # Count testable items = inputs + outputs (excluding ignored)
        testable = [i for i in self._inputs + self._outputs
                    if i["uuid"] not in self._ignored]
        total = len(testable)
        if total == 0:
            self._prog_lbl.config(text="No items loaded")
            self._prog_bar["value"] = 0
            return
        tested = sum(1 for i in testable if self._status_map.get(i["uuid"]) is not None)
        ok  = sum(1 for i in testable if self._status_map.get(i["uuid"]) == "ok")
        nok = sum(1 for i in testable if self._status_map.get(i["uuid"]) == "nok")
        ign = len(self._inputs + self._outputs) - total
        pct = int(tested / total * 100)
        ign_str = f"  ⊘{ign}" if ign else ""
        self._prog_lbl.config(
            text=f"{tested}/{total} tested ({pct}%)   ✓{ok}  ✗{nok}{ign_str}")
        self._prog_bar["value"] = pct

    # ─────────────────────────────────────────────────────────────────────────
    #  SESSION SAVE / LOAD
    # ─────────────────────────────────────────────────────────────────────────
    def _save_session(self):
        if not self._structure:
            messagebox.showwarning("Nothing to save", "Connect and load a project first.")
            return
        mi = self._ms_info
        default_name = (
            f"session_{mi.get('projectName','loxone').replace(' ','_')}"
            f"_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
        )
        path = filedialog.asksaveasfilename(
            title="Save session", defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON session","*.json"),("All","*.*")])
        if not path: return

        testable = [i for i in self._inputs + self._outputs
                    if i["uuid"] not in self._ignored]
        session = {
            "saved_at":   datetime.now().isoformat(),
            "version":    "2.1",
            "ms_info":    self._ms_info,
            "conn_host":  self._conn.get("host",""),
            "statuses":   self._status_map,
            "notes":      self._notes,
            "ignored":    list(self._ignored),
            "summary": {
                "total":    len(testable),
                "ignored":  len(self._ignored),
                "ok":       sum(1 for i in testable if self._status_map.get(i["uuid"])=="ok"),
                "nok":      sum(1 for i in testable if self._status_map.get(i["uuid"])=="nok"),
                "skip":     sum(1 for i in testable if self._status_map.get(i["uuid"])=="skip"),
                "untested": sum(1 for i in testable if self._status_map.get(i["uuid"]) is None),
            }
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(session, f, indent=2, ensure_ascii=False)
            self._log_msg(f"Session saved: {path}", "success")
            messagebox.showinfo("Saved", f"Session saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    def _load_session(self):
        path = filedialog.askopenfilename(
            title="Load session", filetypes=[("JSON session","*.json"),("All","*.*")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                session = json.load(f)
            loaded_status  = session.get("statuses", {})
            loaded_notes   = session.get("notes",    {})
            loaded_ignored = set(session.get("ignored", []))
            self._status_map.update(loaded_status)
            self._notes.update(loaded_notes)
            self._ignored.update(loaded_ignored)

            if not self._structure:
                # Show info about the session but no structure is loaded yet
                mi = session.get("ms_info", {})
                s  = session.get("summary", {})
                messagebox.showinfo("Session loaded (no structure)",
                    f"Session from: {session.get('saved_at','?')}\n"
                    f"Project:  {mi.get('projectName','?')}\n"
                    f"Statuses: ✓{s.get('ok',0)}  ✗{s.get('nok',0)}  "
                    f"→{s.get('skip',0)}  ○{s.get('untested',0)}\n\n"
                    f"Connect to the Miniserver to apply the session.")
            else:
                self._populate_tree()
                self._update_progress()
                s = session.get("summary",{})
                self._log_msg(
                    f"Session loaded: ✓{s.get('ok',0)} ok  "
                    f"✗{s.get('nok',0)} not ok  "
                    f"→{s.get('skip',0)} skip", "success")
                messagebox.showinfo("Session loaded",
                    f"Statuses restored from:\n{path}")
        except Exception as e:
            messagebox.showerror("Load failed", str(e))

    # ─────────────────────────────────────────────────────────────────────────
    #  REPORT
    # ─────────────────────────────────────────────────────────────────────────
    def _populate_room_list(self):
        self._room_list.delete(0,"end")
        rooms = sorted({i["room"] for i in self._all_items})
        for r in rooms: self._room_list.insert("end", r)
        self._room_list.select_set(0,"end")

    def _rooms_all(self):
        self._room_list.select_set(0,"end"); self._update_preview()
    def _rooms_none(self):
        self._room_list.selection_clear(0,"end"); self._update_preview()

    def _selected_rooms(self):
        return {self._room_list.get(i) for i in self._room_list.curselection()}

    def _status_filter_fn(self):
        sv = self._rpt_status.get()
        if sv == "all":   return lambda s: True
        if sv == "ok":    return lambda s: s == "ok"
        if sv == "nok":   return lambda s: s == "nok"
        if sv == "none":  return lambda s: s is None
        return lambda s: True

    def _update_preview(self):
        if not self._structure:
            self._preview_lbl.config(text="— not connected —"); return
        rooms  = self._selected_rooms()
        sfn    = self._status_filter_fn()
        inc_i  = self._inc_inputs.get()
        inc_o  = self._inc_outputs.get()
        inc_x  = self._inc_other.get()

        def _cnt(items, inc):
            if not inc: return 0
            return sum(1 for i in items
                       if i["uuid"] not in self._ignored
                       and i["room"] in rooms
                       and sfn(self._status_map.get(i["uuid"])))

        ci = _cnt(self._inputs,  inc_i)
        co = _cnt(self._outputs, inc_o)
        cx = _cnt(self._others,  inc_x)
        parts = []
        if inc_i: parts.append(f"{ci} inputs")
        if inc_o: parts.append(f"{co} outputs")
        if inc_x: parts.append(f"{cx} other")
        self._preview_lbl.config(
            text=f"{len(rooms)} room(s)  ·  {',  '.join(parts) or 'nothing'}  ·  {ci+co+cx} items total")

    def _generate(self, mode="full"):
        if not self._structure: return
        out = self._out_path.get().strip()
        if not out:
            messagebox.showwarning("No path","Choose a save path first."); return

        sfn   = self._status_filter_fn()
        rooms = self._selected_rooms() if mode == "filtered" else None
        inc_o = self._inc_other.get()

        def _flt(items, inc):
            if not inc: return []
            r = [i for i in items if i["uuid"] not in self._ignored]   # exclude ignored
            if rooms: r = [i for i in r if i["room"] in rooms]
            r = [i for i in r if sfn(self._status_map.get(i["uuid"]))]
            return r

        if mode == "full":
            inp  = _flt(self._inputs,  self._inc_inputs.get())
            outp = _flt(self._outputs, self._inc_outputs.get())
            oth  = _flt(self._others,  inc_o)
        else:
            if not rooms:
                messagebox.showwarning("No rooms","Select at least one room."); return
            inp  = _flt(self._inputs,  self._inc_inputs.get())
            outp = _flt(self._outputs, self._inc_outputs.get())
            oth  = _flt(self._others,  inc_o)

        # Attach status and note to each item for the PDF
        for item in inp + outp + oth:
            item["_status"] = self._status_map.get(item["uuid"])
            item["_note"]   = self._notes.get(item["uuid"], "")

        company = self._company_name.get().strip()
        if self._company_sub.get().strip():
            company = f"{company}\n{self._company_sub.get().strip()}" if company \
                      else self._company_sub.get().strip()
        self._run_thread(self._do_generate, inp, outp, oth, bool(oth), out, company)

    def _do_generate(self, inp, outp, oth, inc_other, out_path, company=""):
        self._set_status("Generating PDF…")
        self._log_msg(f"Generating → {out_path}", "info")
        try:
            generate_pdf(inp, outp, oth, self._ms_info, out_path,
                         include_other=inc_other, company=company)
            self._log_msg(f"PDF saved: {out_path}", "success")
            self._set_status("PDF saved.")
            self.after(0, self._ask_open, out_path)
        except Exception as e:
            self.after(0, self._on_err, str(e))

    def _ask_open(self, path):
        self._busy = False
        if messagebox.askyesno("Done", f"PDF saved:\n{path}\n\nOpen it now?"):
            os.startfile(path)

    # ─────────────────────────────────────────────────────────────────────────
    #  CONFIG  (company / branding)
    # ─────────────────────────────────────────────────────────────────────────
    def _config_path(self):
        """Config file lives next to the EXE (or script when running from source)."""
        if getattr(sys, "frozen", False):
            # Running as a PyInstaller-frozen EXE — write config beside the .exe
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, CONFIG_FILE)

    def _load_config(self):
        try:
            with open(self._config_path(), "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self._company_name.set(cfg.get("company_name", ""))
            self._company_sub .set(cfg.get("company_sub",  ""))
        except (FileNotFoundError, json.JSONDecodeError):
            pass   # first run — no config yet
        self._update_hdr_company()

    def _save_config(self):
        cfg = {
            "company_name": self._company_name.get().strip(),
            "company_sub":  self._company_sub.get().strip(),
        }
        try:
            with open(self._config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
            self._log_msg("Branding saved to config file.", "success")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    def _update_hdr_company(self):
        name = self._company_name.get().strip()
        sub  = self._company_sub.get().strip()
        self._hdr_company.config(text=name)
        # include sub only if different from name and non-empty
        if sub and sub != name:
            self._hdr_company.config(text=f"{name}  —  {sub}" if name else sub)

    # ─────────────────────────────────────────────────────────────────────────
    #  HELPERS
    # ─────────────────────────────────────────────────────────────────────────
    @staticmethod
    def _norm_uuid(u):
        """LoxAPP3.json stores state UUIDs as 8-4-4-16 (3 dashes).
        The HTTP API expects standard 8-4-4-4-12 (4 dashes).
        Example: '1ce71bbb-021d-6281-ffff388ed9c688d5'
              →  '1ce71bbb-021d-6281-ffff-388ed9c688d5'
        """
        parts = u.split('-')
        if len(parts) == 4 and len(parts[3]) == 16:
            last = parts[3]
            return '-'.join(parts[:3] + [last[:4], last[4:]])
        return u   # already 4-dash or unknown format

    def _run_thread(self, fn, *args):
        self._busy = True
        threading.Thread(target=fn, args=args, daemon=True).start()

    def _set_status(self, msg):
        self.after(0, self._status.set, msg)

    def _set_info(self, text):
        def _do():
            self._info.config(state="normal")
            self._info.delete("1.0","end")
            self._info.insert("end", text)
            self._info.config(state="disabled")
        self.after(0, _do)

    def _log_msg(self, msg, tag=None):
        def _do():
            self._log.config(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            self._log.insert("end", f"[{ts}]  {msg}\n")
            if tag:
                s = self._log.index("end - 1 line linestart")
                e = self._log.index("end - 1 char")
                self._log.tag_add(tag, s, e)
            self._log.see("end")
            self._log.config(state="disabled")
        self.after(0, _do)

    def _cmd_log_msg(self, msg, tag=None):
        self._cmd_log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self._cmd_log.insert("end", f"[{ts}]  {msg}\n")
        if tag:
            s = self._cmd_log.index("end - 1 line linestart")
            e = self._cmd_log.index("end - 1 char")
            self._cmd_log.tag_add(tag, s, e)
        self._cmd_log.see("end")
        self._cmd_log.config(state="disabled")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0","end")
        self._log.config(state="disabled")

    def _setup_log_tags(self):
        for lg in (self._log, self._cmd_log):
            lg.tag_config("error",   foreground="#f48771")
            lg.tag_config("success", foreground="#b5cea8")
            lg.tag_config("info",    foreground="#9cdcfe")

    def _center(self):
        self.update_idletasks()
        w,h   = self.winfo_width(), self.winfo_height()
        sw,sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


def main():
    app = App()
    app._setup_log_tags()
    app.mainloop()

if __name__ == "__main__":
    main()
