"""
Excel AI Studio — Main GUI
===========================
Beautiful tkinter GUI with two AI agents:
  • Agent 1: Data Cleaning
  • Agent 2: Power BI Report Generator

Run: python app.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import os
import sys
import shutil
import webbrowser
from pathlib import Path
from datetime import datetime


# ── Resolve paths ──────────────────────────────────────────
BASE = Path(__file__).parent
sys.path.insert(0, str(BASE))

OUTPUT_DIR = BASE / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════════
# COLOUR PALETTE
# ══════════════════════════════════════════════════════════════
C = {
    "bg":        "#0d1117",
    "surface":   "#161b22",
    "surface2":  "#21262d",
    "border":    "#30363d",
    "accent":    "#58a6ff",
    "green":     "#3fb950",
    "orange":    "#d29922",
    "red":       "#f85149",
    "purple":    "#bc8cff",
    "text":      "#e6edf3",
    "muted":     "#8b949e",
    "btn_bg":    "#21262d",
    "btn_hover": "#30363d",
}


def blend(hex_color: str, bg: str = "#161b22", alpha: float = 0.2) -> str:
    """
    Simulate alpha blending: mix hex_color onto bg at given alpha.
    Returns a solid 6-digit hex — tkinter compatible.
    """
    def parse(h):
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    try:
        fr, fg_, fb = parse(hex_color)
        br, bg_, bb = parse(bg)
        r = int(fr * alpha + br * (1 - alpha))
        g = int(fg_ * alpha + bg_ * (1 - alpha))
        b = int(fb * alpha + bb * (1 - alpha))
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return bg

FONTS = {
    "title":  ("Consolas", 14, "bold"),
    "label":  ("Consolas", 10),
    "small":  ("Consolas", 9),
    "log":    ("Consolas", 9),
    "h1":     ("Consolas", 18, "bold"),
    "badge":  ("Consolas", 8, "bold"),
}


# ══════════════════════════════════════════════════════════════
# HELPER WIDGETS
# ══════════════════════════════════════════════════════════════
class StyledButton(tk.Button):
    def __init__(self, parent, text, command=None, accent=False, **kw):
        bg = C["accent"] if accent else C["btn_bg"]
        fg = C["bg"]     if accent else C["text"]
        super().__init__(
            parent, text=text, command=command,
            bg=bg, fg=fg, relief="flat",
            font=FONTS["label"], cursor="hand2",
            padx=14, pady=6, **kw
        )
        self.accent = accent
        self._bg = bg
        self.bind("<Enter>",  self._on_enter)
        self.bind("<Leave>",  self._on_leave)

    def _on_enter(self, e):
        self.config(bg=C["btn_hover"] if not self.accent else "#79b8ff")

    def _on_leave(self, e):
        self.config(bg=self._bg)


class Badge(tk.Label):
    def __init__(self, parent, text, color=None, **kw):
        color = color or C["accent"]
        super().__init__(
            parent, text=f" {text} ",
            bg=blend(color, C["surface"], 0.18), fg=color,
            font=FONTS["badge"], padx=4, pady=1, **kw
        )


class SectionCard(tk.Frame):
    def __init__(self, parent, title="", **kw):
        super().__init__(parent, bg=C["surface"],
                         highlightbackground=C["border"],
                         highlightthickness=1, **kw)
        if title:
            tk.Label(self, text=title, bg=C["surface"],
                     fg=C["muted"], font=FONTS["small"],
                     pady=6, padx=12).pack(anchor="w")
            ttk.Separator(self, orient="horizontal").pack(fill="x", padx=8)


# ══════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ══════════════════════════════════════════════════════════════
class ExcelAIStudio(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Excel AI Studio")
        self.geometry("1100x780")
        self.minsize(900, 640)
        self.configure(bg=C["bg"])
        self.resizable(True, True)

        # State
        self.input_file   = tk.StringVar()
        self.clean_file   = tk.StringVar()
        self.agent1_running = False
        self.agent2_running = False
        self.profile_data   = {}

        self._build_ui()
        self._log_agent1("Welcome to Excel AI Studio", "info")
        self._log_agent1("Step 1: Select your Excel file to begin.", "muted")

    # ── UI CONSTRUCTION ──────────────────────────────────────

    def _build_ui(self):
        # Top bar
        topbar = tk.Frame(self, bg=C["surface"],
                          highlightbackground=C["border"],
                          highlightthickness=1, height=48)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)

        tk.Label(topbar, text="⬡  EXCEL AI STUDIO",
                 bg=C["surface"], fg=C["accent"],
                 font=FONTS["title"], padx=16).pack(side="left", pady=10)

        Badge(topbar, "v1.0", C["green"]).pack(side="left", pady=10)

        tk.Label(topbar, text="LangChain + Power BI Agent",
                 bg=C["surface"], fg=C["muted"],
                 font=FONTS["small"]).pack(side="right", padx=16, pady=10)

        # Main notebook (tabs)
        self.nb = ttk.Notebook(self)
        self._style_notebook()
        self.nb.pack(fill="both", expand=True, padx=0, pady=0)

        self._build_tab_home()
        self._build_tab_agent1()
        self._build_tab_agent2()
        self._build_tab_output()

    def _style_notebook(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook",
                        background=C["bg"], borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=C["surface2"],
                        foreground=C["muted"],
                        padding=[16, 8],
                        font=FONTS["label"])
        style.map("TNotebook.Tab",
                  background=[("selected", C["bg"])],
                  foreground=[("selected", C["accent"])])

    # ── TAB 0: HOME ─────────────────────────────────────────

    def _build_tab_home(self):
        tab = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(tab, text="  Home  ")

        # Hero
        hero = tk.Frame(tab, bg=C["bg"])
        hero.pack(fill="both", expand=True, padx=60, pady=40)

        tk.Label(hero, text="Excel AI Studio",
                 bg=C["bg"], fg=C["text"],
                 font=("Consolas", 26, "bold")).pack(anchor="w")

        tk.Label(hero, text="Two AI agents. Clean data. Power BI ready.",
                 bg=C["bg"], fg=C["muted"],
                 font=("Consolas", 12)).pack(anchor="w", pady=(4, 24))

        # Pipeline flow
        flow = tk.Frame(hero, bg=C["bg"])
        flow.pack(fill="x")

        steps = [
            ("📂", "Your Excel File", "Messy / raw data", C["muted"]),
            ("→", "", "", C["border"]),
            ("🤖", "Agent 1", "Clean & profile", C["accent"]),
            ("→", "", "", C["border"]),
            ("✨", "Clean Excel", "Report-ready data", C["green"]),
            ("→", "", "", C["border"]),
            ("🤖", "Agent 2", "Power BI setup", C["purple"]),
            ("→", "", "", C["border"]),
            ("📊", "Power BI", "Live dashboard", C["orange"]),
        ]

        for icon, title, sub, color in steps:
            if icon == "→":
                tk.Label(flow, text="───►", bg=C["bg"],
                         fg=C["border"], font=("Consolas", 14)).pack(
                    side="left", padx=4)
                continue
            card = tk.Frame(flow, bg=C["surface"],
                            highlightbackground=color,
                            highlightthickness=1, padx=14, pady=10)
            card.pack(side="left")
            tk.Label(card, text=icon, bg=C["surface"],
                     font=("Consolas", 20)).pack()
            tk.Label(card, text=title, bg=C["surface"],
                     fg=color, font=FONTS["label"]).pack()
            tk.Label(card, text=sub, bg=C["surface"],
                     fg=C["muted"], font=FONTS["small"]).pack()

        # File picker
        picker_frame = tk.Frame(hero, bg=C["surface"],
                                highlightbackground=C["border"],
                                highlightthickness=1, pady=20, padx=20)
        picker_frame.pack(fill="x", pady=30)

        tk.Label(picker_frame, text="Select your Excel file to get started",
                 bg=C["surface"], fg=C["text"],
                 font=FONTS["label"]).pack(anchor="w")

        row = tk.Frame(picker_frame, bg=C["surface"])
        row.pack(fill="x", pady=8)

        tk.Entry(row, textvariable=self.input_file,
                 bg=C["surface2"], fg=C["text"],
                 insertbackground=C["text"],
                 relief="flat", font=FONTS["label"],
                 highlightbackground=C["border"],
                 highlightthickness=1).pack(
            side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        StyledButton(row, "Browse...", command=self._browse_file).pack(side="left")
        StyledButton(row, "Use Sample Data",
                     command=self._use_sample).pack(side="left", padx=(8, 0))
        StyledButton(row, "Start Agent 1 →",
                     command=lambda: self.nb.select(1),
                     accent=True).pack(side="right")

        # Features
        feat_row = tk.Frame(hero, bg=C["bg"])
        feat_row.pack(fill="x", pady=(0, 20))

        features = [
            (C["accent"],  "Agent 1 — Data Cleaner",
             "Auto-detects 10+ issue types.\nFixes nulls, dates, duplicates,\ncasing, categories, numerics."),
            (C["purple"],  "Agent 2 — Power BI Builder",
             "Generates DAX measures,\nPower Query M code,\nstep-by-step setup guide."),
            (C["green"],   "LangChain Tools",
             "10 specialised tools per agent.\nEach tool does one job well.\nChained for full automation."),
        ]
        for color, title, desc in features:
            f = tk.Frame(feat_row, bg=C["surface"],
                         highlightbackground=blend(color, C["surface"], 0.5),
                         highlightthickness=1, padx=16, pady=14)
            f.pack(side="left", fill="both", expand=True,
                   padx=(0, 12))
            tk.Label(f, text=title, bg=C["surface"],
                     fg=color, font=FONTS["label"]).pack(anchor="w")
            tk.Label(f, text=desc, bg=C["surface"],
                     fg=C["muted"], font=FONTS["small"],
                     justify="left").pack(anchor="w", pady=(6, 0))

    # ── TAB 1: AGENT 1 — DATA CLEANING ─────────────────────

    def _build_tab_agent1(self):
        tab = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(tab, text="  Agent 1: Clean Data  ")

        # Left panel
        left = tk.Frame(tab, bg=C["bg"], width=340)
        left.pack(side="left", fill="y", padx=(12, 0), pady=12)
        left.pack_propagate(False)

        tk.Label(left, text="🤖  Data Cleaning Agent",
                 bg=C["bg"], fg=C["accent"],
                 font=FONTS["title"]).pack(anchor="w", pady=(0, 4))
        tk.Label(left, text="10 LangChain tools · Auto-detects & fixes all issues",
                 bg=C["bg"], fg=C["muted"],
                 font=FONTS["small"]).pack(anchor="w", pady=(0, 12))

        # File row
        file_card = SectionCard(left, "Input File")
        file_card.pack(fill="x", pady=(0, 8))

        row = tk.Frame(file_card, bg=C["surface"])
        row.pack(fill="x", padx=8, pady=8)
        tk.Entry(row, textvariable=self.input_file,
                 bg=C["surface2"], fg=C["text"],
                 relief="flat", font=FONTS["small"],
                 highlightbackground=C["border"],
                 highlightthickness=1,
                 state="readonly").pack(
            side="left", fill="x", expand=True, ipady=4)
        StyledButton(row, "...", command=self._browse_file,
                     padx=8, pady=4).pack(side="left", padx=(4, 0))

        # Tools list
        tools_card = SectionCard(left, "Agent Tools")
        tools_card.pack(fill="x", pady=(0, 8))

        self.tool_labels = {}
        tools = [
            ("profile_data",      "1. Profile Data Quality"),
            ("fix_column_names",  "2. Fix Column Names"),
            ("fix_duplicates",    "3. Remove Duplicates"),
            ("fix_text",          "4. Fix Text & Case"),
            ("fix_dates",         "5. Parse Dates"),
            ("fix_nulls",         "6. Handle Nulls"),
            ("fix_numerics",      "7. Fix Numerics"),
            ("fix_categories",    "8. Standardise Categories"),
            ("add_derived",       "9. Add Derived Columns"),
            ("export_clean",      "10. Export Clean File"),
        ]
        for key, label in tools:
            row = tk.Frame(tools_card, bg=C["surface"])
            row.pack(fill="x", padx=8, pady=2)
            status = tk.Label(row, text="○", bg=C["surface"],
                              fg=C["border"], font=FONTS["small"], width=2)
            status.pack(side="left")
            tk.Label(row, text=label, bg=C["surface"],
                     fg=C["muted"], font=FONTS["small"]).pack(side="left")
            self.tool_labels[key] = status

        # Progress
        prog_card = SectionCard(left, "Progress")
        prog_card.pack(fill="x", pady=(0, 8))

        self.prog_var = tk.DoubleVar(value=0)
        self.prog_bar = ttk.Progressbar(prog_card,
                                        variable=self.prog_var,
                                        maximum=100)
        self.prog_bar.pack(fill="x", padx=8, pady=8)
        self.prog_label = tk.Label(prog_card, text="Ready",
                                   bg=C["surface"], fg=C["muted"],
                                   font=FONTS["small"])
        self.prog_label.pack(anchor="w", padx=8, pady=(0, 8))

        # Buttons
        btn_row = tk.Frame(left, bg=C["bg"])
        btn_row.pack(fill="x", pady=4)
        self.btn_clean = StyledButton(btn_row, "▶  Run Cleaning Agent",
                                      command=self._run_agent1,
                                      accent=True)
        self.btn_clean.pack(fill="x", pady=(0, 4))

        StyledButton(btn_row, "Use Sample Data",
                     command=self._use_sample).pack(fill="x")

        # Right panel — log
        right = tk.Frame(tab, bg=C["bg"])
        right.pack(side="right", fill="both", expand=True,
                   padx=12, pady=12)

        tk.Label(right, text="Agent Log",
                 bg=C["bg"], fg=C["muted"],
                 font=FONTS["small"]).pack(anchor="w")

        self.log1 = scrolledtext.ScrolledText(
            right, bg=C["surface"], fg=C["text"],
            font=FONTS["log"], relief="flat",
            insertbackground=C["text"],
            highlightbackground=C["border"],
            highlightthickness=1,
            wrap="word", state="disabled"
        )
        self.log1.pack(fill="both", expand=True, pady=(4, 0))
        self.log1.tag_config("info",  foreground=C["accent"])
        self.log1.tag_config("ok",    foreground=C["green"])
        self.log1.tag_config("warn",  foreground=C["orange"])
        self.log1.tag_config("error", foreground=C["red"])
        self.log1.tag_config("muted", foreground=C["muted"])
        self.log1.tag_config("tool",  foreground=C["purple"])

        # Quality scorecard
        self.quality_frame = tk.Frame(right, bg=C["surface"],
                                      highlightbackground=C["border"],
                                      highlightthickness=1)
        self.quality_frame.pack(fill="x", pady=(8, 0))
        tk.Label(self.quality_frame,
                 text="Quality Report will appear here after cleaning",
                 bg=C["surface"], fg=C["muted"],
                 font=FONTS["small"], pady=10).pack()

    # ── TAB 2: AGENT 2 — POWER BI ──────────────────────────

    def _build_tab_agent2(self):
        tab = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(tab, text="  Agent 2: Power BI  ")

        left = tk.Frame(tab, bg=C["bg"], width=340)
        left.pack(side="left", fill="y", padx=(12, 0), pady=12)
        left.pack_propagate(False)

        tk.Label(left, text="📊  Power BI Agent",
                 bg=C["bg"], fg=C["purple"],
                 font=FONTS["title"]).pack(anchor="w", pady=(0, 4))
        tk.Label(left,
                 text="Generates DAX, Power Query & setup guide",
                 bg=C["bg"], fg=C["muted"],
                 font=FONTS["small"]).pack(anchor="w", pady=(0, 12))

        # Clean file
        cf_card = SectionCard(left, "Clean File (from Agent 1)")
        cf_card.pack(fill="x", pady=(0, 8))
        row = tk.Frame(cf_card, bg=C["surface"])
        row.pack(fill="x", padx=8, pady=8)
        tk.Entry(row, textvariable=self.clean_file,
                 bg=C["surface2"], fg=C["text"],
                 relief="flat", font=FONTS["small"],
                 highlightbackground=C["border"],
                 highlightthickness=1,
                 state="readonly").pack(
            side="left", fill="x", expand=True, ipady=4)

        # Tools list
        tools2_card = SectionCard(left, "Agent Tools")
        tools2_card.pack(fill="x", pady=(0, 8))

        self.tool2_labels = {}
        tools2 = [
            ("schema",    "1. Analyse Schema"),
            ("dax",       "2. Generate DAX Measures"),
            ("pq",        "3. Generate Power Query"),
            ("guide",     "4. Generate Setup Guide"),
            ("preview",   "5. Build Report Preview"),
        ]
        for key, label in tools2:
            row = tk.Frame(tools2_card, bg=C["surface"])
            row.pack(fill="x", padx=8, pady=2)
            st = tk.Label(row, text="○", bg=C["surface"],
                          fg=C["border"], font=FONTS["small"], width=2)
            st.pack(side="left")
            tk.Label(row, text=label, bg=C["surface"],
                     fg=C["muted"], font=FONTS["small"]).pack(side="left")
            self.tool2_labels[key] = st

        # Progress
        prog2_card = SectionCard(left, "Progress")
        prog2_card.pack(fill="x", pady=(0, 8))
        self.prog2_var = tk.DoubleVar(value=0)
        ttk.Progressbar(prog2_card, variable=self.prog2_var,
                        maximum=100).pack(fill="x", padx=8, pady=8)
        self.prog2_label = tk.Label(prog2_card, text="Run Agent 1 first",
                                    bg=C["surface"], fg=C["muted"],
                                    font=FONTS["small"])
        self.prog2_label.pack(anchor="w", padx=8, pady=(0, 8))

        # Output files
        out_card = SectionCard(left, "Output Files")
        out_card.pack(fill="x", pady=(0, 8))
        self.out_links = tk.Frame(out_card, bg=C["surface"])
        self.out_links.pack(fill="x", padx=8, pady=8)
        tk.Label(self.out_links, text="Will appear after running agent",
                 bg=C["surface"], fg=C["muted"],
                 font=FONTS["small"]).pack()

        self.btn_pbi = StyledButton(left, "▶  Run Power BI Agent",
                                    command=self._run_agent2, accent=True)
        self.btn_pbi.pack(fill="x", pady=4)

        # Right
        right = tk.Frame(tab, bg=C["bg"])
        right.pack(side="right", fill="both", expand=True,
                   padx=12, pady=12)
        tk.Label(right, text="Agent Log",
                 bg=C["bg"], fg=C["muted"],
                 font=FONTS["small"]).pack(anchor="w")
        self.log2 = scrolledtext.ScrolledText(
            right, bg=C["surface"], fg=C["text"],
            font=FONTS["log"], relief="flat",
            insertbackground=C["text"],
            highlightbackground=C["border"],
            highlightthickness=1,
            wrap="word", state="disabled"
        )
        self.log2.pack(fill="both", expand=True, pady=(4, 0))
        self.log2.tag_config("info",  foreground=C["purple"])
        self.log2.tag_config("ok",    foreground=C["green"])
        self.log2.tag_config("warn",  foreground=C["orange"])
        self.log2.tag_config("error", foreground=C["red"])
        self.log2.tag_config("muted", foreground=C["muted"])
        self.log2.tag_config("tool",  foreground=C["accent"])

    # ── TAB 3: OUTPUT ────────────────────────────────────────

    def _build_tab_output(self):
        tab = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(tab, text="  Output Files  ")

        tk.Label(tab, text="Generated Files",
                 bg=C["bg"], fg=C["text"],
                 font=FONTS["title"]).pack(anchor="w", padx=16, pady=(16, 4))

        tk.Label(tab,
                 text="All files generated by the agents appear here.",
                 bg=C["bg"], fg=C["muted"],
                 font=FONTS["small"]).pack(anchor="w", padx=16)

        self.output_frame = tk.Frame(tab, bg=C["bg"])
        self.output_frame.pack(fill="both", expand=True, padx=16, pady=12)

        StyledButton(tab, "🔄 Refresh", command=self._refresh_output).pack(
            anchor="w", padx=16, pady=(0, 8))

        self._refresh_output()

    # ── ACTIONS ──────────────────────────────────────────────

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All", "*.*")]
        )
        if path:
            self.input_file.set(path)
            self._log_agent1(f"File selected: {Path(path).name}", "info")

    def _use_sample(self):
        sample = BASE / "data" / "sample_messy.xlsx"
        if not sample.exists():
            messagebox.showerror("Error", "Sample file not found.\nRun: python agents/cleaning_agent.py")
            return
        self.input_file.set(str(sample))
        self._log_agent1("Sample data loaded: sample_messy.xlsx", "info")
        self._log_agent1("This file has: nulls, duplicates, mixed case,", "muted")
        self._log_agent1("inconsistent dates, negative values, string nulls.", "muted")
        self.nb.select(1)

    def _set_tool_status(self, tool_dict, key, status):
        """Update tool status icon in the sidebar."""
        icons = {"pending": ("○", C["border"]),
                 "running": ("◉", C["orange"]),
                 "done":    ("●", C["green"]),
                 "error":   ("✗", C["red"])}
        icon, color = icons.get(status, ("○", C["border"]))
        if key in tool_dict:
            tool_dict[key].config(text=icon, fg=color)

    def _run_agent1(self):
        if self.agent1_running:
            return
        path = self.input_file.get().strip()
        if not path or not Path(path).exists():
            messagebox.showerror("No File", "Please select an Excel file first.")
            return

        self.agent1_running = True
        self.btn_clean.config(state="disabled", text="Running...")
        for k in self.tool_labels:
            self._set_tool_status(self.tool_labels, k, "pending")

        thread = threading.Thread(
            target=self._agent1_worker, args=(path,), daemon=True)
        thread.start()

    def _agent1_worker(self, path):
        try:
            from agents.cleaning_agent import run_cleaning_pipeline

            output_path = str(OUTPUT_DIR / "clean_data.xlsx")

            self._log_agent1("", "muted")
            self._log_agent1("═" * 52, "muted")
            self._log_agent1("  🤖 DATA CLEANING AGENT STARTING", "info")
            self._log_agent1(f"  Input: {Path(path).name}", "muted")
            self._log_agent1("═" * 52, "muted")

            tool_keys = list(self.tool_labels.keys())
            tool_idx  = [0]

            def progress_cb(step, total, message):
                pct = int(step / total * 100)
                self.after(0, lambda: self.prog_var.set(pct))
                self.after(0, lambda: self.prog_label.config(text=message))
                self._log_agent1(f"  [{step:02d}/{total}] {message}", "tool")

                # Advance tool icons
                idx = tool_idx[0]
                if idx < len(tool_keys):
                    self.after(0, lambda k=tool_keys[idx]:
                               self._set_tool_status(self.tool_labels, k, "running"))
                if idx > 0 and (idx-1) < len(tool_keys):
                    self.after(0, lambda k=tool_keys[idx-1]:
                               self._set_tool_status(self.tool_labels, k, "done"))
                tool_idx[0] += 1

            result = run_cleaning_pipeline(path, output_path, progress_cb)

            # Mark all done
            for k in self.tool_labels:
                self.after(0, lambda kk=k:
                           self._set_tool_status(self.tool_labels, kk, "done"))

            imp = result.get("improvement", {})
            self._log_agent1("", "muted")
            self._log_agent1("═" * 52, "ok")
            self._log_agent1("  ✅ CLEANING COMPLETE!", "ok")
            self._log_agent1(f"  Quality Score:  {imp.get('quality_before',0)} → {imp.get('quality_after',0)}", "ok")
            self._log_agent1(f"  Issues Fixed:   {imp.get('issues_before',0)} → {imp.get('issues_after',0)}", "ok")
            self._log_agent1(f"  Rows:           {imp.get('rows_before',0)} → {imp.get('rows_after',0)}", "ok")
            self._log_agent1(f"  Output:         clean_data.xlsx", "ok")
            self._log_agent1("═" * 52, "ok")

            self.clean_file.set(output_path)
            self.prog_var.set(100)
            self.prog_label.config(text="✅ Complete!")

            self.after(0, lambda r=result: self._show_quality_report(r))
            self.after(0, self._refresh_output)

        except Exception as e:
            self._log_agent1(f"ERROR: {e}", "error")
            import traceback
            self._log_agent1(traceback.format_exc(), "error")
        finally:
            self.agent1_running = False
            self.after(0, lambda: self.btn_clean.config(
                state="normal", text="▶  Run Cleaning Agent"))

    def _show_quality_report(self, result):
        """Show before/after quality scorecard."""
        for w in self.quality_frame.winfo_children():
            w.destroy()

        imp = result.get("improvement", {})

        tk.Label(self.quality_frame, text="Quality Report",
                 bg=C["surface"], fg=C["muted"],
                 font=FONTS["small"], pady=6, padx=12).pack(anchor="w")

        row = tk.Frame(self.quality_frame, bg=C["surface"])
        row.pack(fill="x", padx=12, pady=(0, 8))

        metrics = [
            ("Quality Score", imp.get("quality_before",0),
             imp.get("quality_after",0), "%"),
            ("Issues",        imp.get("issues_before",0),
             imp.get("issues_after",0), ""),
            ("Rows",          imp.get("rows_before",0),
             imp.get("rows_after",0), ""),
        ]
        for label, before, after, unit in metrics:
            col = tk.Frame(row, bg=C["surface"])
            col.pack(side="left", expand=True)
            tk.Label(col, text=label, bg=C["surface"],
                     fg=C["muted"], font=FONTS["small"]).pack()
            tk.Label(col, text=f"{before}{unit} → {after}{unit}",
                     bg=C["surface"], fg=C["green"],
                     font=FONTS["label"]).pack()

        StyledButton(self.quality_frame,
                     "→ Go to Power BI Agent",
                     command=lambda: self.nb.select(2),
                     accent=True).pack(padx=12, pady=(0, 8))

    def _run_agent2(self):
        if self.agent2_running:
            return
        path = self.clean_file.get().strip()
        if not path or not Path(path).exists():
            messagebox.showinfo(
                "Run Agent 1 First",
                "Please run the Data Cleaning Agent first\n"
                "to generate the clean Excel file.")
            self.nb.select(1)
            return

        self.agent2_running = True
        self.btn_pbi.config(state="disabled", text="Running...")
        for k in self.tool2_labels:
            self._set_tool_status(self.tool2_labels, k, "pending")

        thread = threading.Thread(
            target=self._agent2_worker, args=(path,), daemon=True)
        thread.start()

    def _agent2_worker(self, path):
        try:
            from agents.powerbi_agent import run_powerbi_pipeline

            out_dir = str(OUTPUT_DIR)

            self._log_agent2("")
            self._log_agent2("═" * 52, "muted")
            self._log_agent2("  📊 POWER BI AGENT STARTING", "info")
            self._log_agent2(f"  Input: {Path(path).name}", "muted")
            self._log_agent2("═" * 52, "muted")

            tool_map = {
                1: "schema", 2: "dax", 3: "pq", 4: "guide", 5: "preview"
            }

            def progress_cb(step, total, message):
                pct = int(step / total * 100)
                self.after(0, lambda: self.prog2_var.set(pct))
                self.after(0, lambda: self.prog2_label.config(text=message))
                self._log_agent2(f"  [{step:02d}/{total}] {message}", "tool")
                key = tool_map.get(step)
                if key:
                    self.after(0, lambda k=key:
                               self._set_tool_status(self.tool2_labels, k, "running"))
                prev_key = tool_map.get(step - 1)
                if prev_key:
                    self.after(0, lambda k=prev_key:
                               self._set_tool_status(self.tool2_labels, k, "done"))

            result = run_powerbi_pipeline(path, out_dir, progress_cb)

            for k in self.tool2_labels:
                self.after(0, lambda kk=k:
                           self._set_tool_status(self.tool2_labels, kk, "done"))

            self._log_agent2("")
            self._log_agent2("═" * 52, "ok")
            self._log_agent2("  ✅ POWER BI FILES READY!", "ok")

            files = result.get("files_created", [])
            for f in files:
                if f and Path(f).exists():
                    self._log_agent2(f"  📄 {Path(f).name}", "ok")

            self._log_agent2(f"  📐 DAX Measures: {result.get('dax_measure_count',0)}", "ok")
            steps = result.get("power_query_steps", [])
            self._log_agent2(f"  ⚡ Power Query Steps: {len(steps)}", "ok")
            self._log_agent2("═" * 52, "ok")
            self._log_agent2("", "muted")
            self._log_agent2("  NEXT STEPS:", "info")
            self._log_agent2("  1. Open Power BI Desktop", "muted")
            self._log_agent2("  2. Load clean_data.xlsx", "muted")
            self._log_agent2("  3. Paste Power Query code", "muted")
            self._log_agent2("  4. Add DAX measures", "muted")
            self._log_agent2("  5. Follow setup guide", "muted")

            self.prog2_var.set(100)
            self.prog2_label.config(text="✅ Complete!")
            self.after(0, self._refresh_output)
            self.after(0, self._update_output_links)

        except Exception as e:
            self._log_agent2(f"ERROR: {e}", "error")
            import traceback
            self._log_agent2(traceback.format_exc(), "error")
        finally:
            self.agent2_running = False
            self.after(0, lambda: self.btn_pbi.config(
                state="normal", text="▶  Run Power BI Agent"))

    def _update_output_links(self):
        for w in self.out_links.winfo_children():
            w.destroy()
        for f in OUTPUT_DIR.iterdir():
            if f.is_file():
                row = tk.Frame(self.out_links, bg=C["surface"])
                row.pack(fill="x", pady=2)
                tk.Label(row, text=f.suffix.upper(),
                         bg=C["surface"], fg=C["purple"],
                         font=FONTS["small"], width=6).pack(side="left")
                btn = tk.Button(row, text=f.name,
                                bg=C["surface"], fg=C["accent"],
                                relief="flat", font=FONTS["small"],
                                cursor="hand2",
                                command=lambda p=f: self._open_file(p))
                btn.pack(side="left")

    def _open_file(self, path):
        import subprocess, platform
        if platform.system() == "Darwin":
            subprocess.run(["open", str(path)])
        elif platform.system() == "Windows":
            os.startfile(str(path))
        else:
            subprocess.run(["xdg-open", str(path)])

    def _refresh_output(self):
        for w in self.output_frame.winfo_children():
            w.destroy()
        files = sorted(OUTPUT_DIR.iterdir()) if OUTPUT_DIR.exists() else []
        if not files:
            tk.Label(self.output_frame,
                     text="No output files yet. Run the agents first.",
                     bg=C["bg"], fg=C["muted"],
                     font=FONTS["small"]).pack(anchor="w", pady=20)
            return

        for f in files:
            if f.is_file():
                row = tk.Frame(self.output_frame, bg=C["surface"],
                               highlightbackground=C["border"],
                               highlightthickness=1)
                row.pack(fill="x", pady=3)

                ext_colors = {".xlsx": C["green"], ".txt": C["accent"],
                              ".m": C["purple"], ".md": C["orange"],
                              ".html": C["red"], ".json": C["muted"]}
                color = ext_colors.get(f.suffix, C["muted"])

                Badge(row, f.suffix.upper()[1:], color).pack(
                    side="left", padx=8, pady=8)
                tk.Label(row, text=f.name, bg=C["surface"],
                         fg=C["text"], font=FONTS["label"]).pack(
                    side="left", padx=4)
                size = f.stat().st_size
                size_str = f"{size/1024:.1f} KB" if size > 1024 else f"{size} B"
                tk.Label(row, text=size_str, bg=C["surface"],
                         fg=C["muted"], font=FONTS["small"]).pack(side="right", padx=8)
                StyledButton(row, "Open",
                             command=lambda p=f: self._open_file(p),
                             padx=8, pady=3).pack(side="right", padx=4)

    # ── LOGGING ──────────────────────────────────────────────

    def _log_agent1(self, msg, tag=""):
        self._log_to(self.log1, msg, tag)

    def _log_agent2(self, msg="", tag=""):
        self._log_to(self.log2, msg, tag)

    def _log_to(self, widget, msg, tag):
        def _do():
            widget.config(state="normal")
            ts = datetime.now().strftime("%H:%M:%S")
            line = f"[{ts}] {msg}\n" if msg else "\n"
            widget.insert("end", line, tag if tag else "")
            widget.see("end")
            widget.config(state="disabled")
        self.after(0, _do)


# ══════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = ExcelAIStudio()
    app.mainloop()
