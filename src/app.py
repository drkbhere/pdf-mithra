#!/usr/bin/env python3
"""Pdf-Mithra — Friend of PDFs."""

import json
import os
import platform
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"

sys.path.insert(0, str(Path(__file__).parent))

from extract import (  # noqa: E402
    extract_all,
    format_markdown,
    format_markdown_with_metadata,
    format_obsidian,
    format_batch_markdown,
    format_csv,
    format_markdown_table,
    format_html,
    COLOR_HEX,
    COLOR_EMOJI,
    TYPE_EMOJI,
    MEANING_PRESETS,
)

try:
    from pynput import keyboard as _kb
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False

# ── Identity ──────────────────────────────────────────────────────────────────
APP_NAME    = "Pdf-Mithra"
APP_TAGLINE = "Your friendly companion for PDF highlights."

# ── Design tokens ─────────────────────────────────────────────────────────────
BG       = "#F6F5F3"   # warm off-white window background
SURFACE  = "#FFFFFF"   # card / panel surface
SURFACE2 = "#EDECEA"   # drop zone, section headers
BORDER   = "#E2DDD8"   # subtle borders
ACCENT   = "#5C5FEF"   # indigo — primary action colour
TEXT     = "#1C1917"   # near-black
TEXT2    = "#6B6560"   # secondary text
TEXT3    = "#A09891"   # muted / placeholder text
GREEN    = "#22C55E"   # success feedback

# ── Platform-specific constants ───────────────────────────────────────────────
if IS_WIN:
    _APP_SUPPORT = Path(os.environ.get("APPDATA", Path.home())) / APP_NAME
    _FONT_UI     = "Segoe UI"
    _FONT_HEAD   = "Segoe UI"
    _FONT_MONO   = "Consolas"
    _CURSOR_HAND = "hand2"
    _MOD         = "Control"
    _MOD_LABEL   = "Ctrl"
else:
    _APP_SUPPORT = Path.home() / "Library" / "Application Support" / APP_NAME
    _FONT_UI     = "SF Pro Text"
    _FONT_HEAD   = "SF Pro Display"
    _FONT_MONO   = "SF Mono"
    _CURSOR_HAND = "pointinghand"
    _MOD         = "Command"
    _MOD_LABEL   = "Cmd"

RECENT_FILES_PATH = _APP_SUPPORT / "recent.json"
MEANINGS_PATH     = _APP_SUPPORT / "color_meanings.json"

MAX_RECENT = 5
OUTPUT_FORMATS = [
    "Standard Markdown",
    "Markdown + Metadata",
    "Obsidian Callouts",
    "Markdown Table",
    "CSV",
    "HTML Report",
]


# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("820x720")
        self.minsize(600, 480)
        self.configure(bg=BG)

        # ── State ─────────────────────────────────────────────────────────────
        self.annotations: list[dict]  = []
        self.pdf_path: str | None     = None
        self.batch_mode: bool         = False
        self.batch_results: list[dict]= []
        self._metadata: dict | None   = None
        self._word_counts: dict | None= None
        self._recent_files            = self._load_recent_files()
        self._raw_mode                = False
        self._hotkey_listener         = None
        self._color_meanings: dict    = self._load_meanings()

        # ── tk vars ───────────────────────────────────────────────────────────
        self._auto_copy     = tk.BooleanVar(value=False)
        self._with_context  = tk.BooleanVar(value=False)
        self._deep_links    = tk.BooleanVar(value=False)
        self._output_format = tk.StringVar(value=OUTPUT_FORMATS[0])
        self._hotkey_on     = tk.BooleanVar(value=False)

        self._build_menu()
        self._build_ui()
        self._try_enable_dnd()

        if IS_MAC:
            self.createcommand("::tk::mac::OpenDocument", self._mac_open_doc)
            self.createcommand("::tk::mac::ReopenApplication", self.deiconify)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf"):
            self.after(100, lambda: self._load_pdfs([sys.argv[1]]))

    # =========================================================================
    # Menu
    # =========================================================================
    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=False)
        _open_viewer_label  = "Open PDF in Preview" if IS_MAC else "Open PDF"
        _reveal_label       = "Reveal in Finder"    if IS_MAC else "Show in Explorer"

        file_menu.add_command(label="Open PDF…",
                              accelerator=f"{_MOD_LABEL}+O", command=self._open_file)
        file_menu.add_separator()
        self._recent_menu = tk.Menu(file_menu, tearoff=False)
        file_menu.add_cascade(label="Open Recent", menu=self._recent_menu)
        self._update_recent_menu()
        file_menu.add_separator()
        file_menu.add_command(label="Save Annotations As…",
                              accelerator=f"{_MOD_LABEL}+S", command=self._save_as)
        file_menu.add_separator()
        file_menu.add_command(label=_open_viewer_label,
                              accelerator=f"{_MOD_LABEL}+Shift+O",
                              command=self._open_in_preview)
        file_menu.add_command(label=_reveal_label, command=self._reveal_in_finder)
        file_menu.add_separator()
        file_menu.add_command(label="Close PDF", command=self._clear_pdf)
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=False)
        edit_menu.add_command(label="Copy Annotations",
                              accelerator=f"{_MOD_LABEL}+Shift+C",
                              command=self._copy_to_clipboard)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        self.config(menu=menubar)

        for key in (f"<{_MOD}-o>", f"<{_MOD}-O>"):
            self.bind_all(key, lambda e: self._open_file())
        for key in (f"<{_MOD}-Shift-c>", f"<{_MOD}-Shift-C>"):
            self.bind_all(key, lambda e: self._copy_to_clipboard())
        for key in (f"<{_MOD}-s>", f"<{_MOD}-S>"):
            self.bind_all(key, lambda e: self._save_as())
        for key in (f"<{_MOD}-Shift-o>", f"<{_MOD}-Shift-O>"):
            self.bind_all(key, lambda e: self._open_in_preview())
        if IS_MAC:
            for key in ("<Command-Alt-e>", "<Command-Alt-E>"):
                self.bind_all(key, lambda e: self._run_hotkey_extraction())

    # =========================================================================
    # Recent files
    # =========================================================================
    def _update_recent_menu(self):
        self._recent_menu.delete(0, "end")
        if not self._recent_files:
            self._recent_menu.add_command(label="No Recent Files", state="disabled")
        else:
            for path in self._recent_files:
                self._recent_menu.add_command(
                    label=Path(path).name,
                    command=lambda p=path: self._load_pdfs([p]),
                )
            self._recent_menu.add_separator()
            self._recent_menu.add_command(label="Clear Recent Files",
                                          command=self._clear_recent)

    def _load_recent_files(self) -> list[str]:
        try:
            if RECENT_FILES_PATH.exists():
                data = json.loads(RECENT_FILES_PATH.read_text())
                return [p for p in data if Path(p).exists()][:MAX_RECENT]
        except Exception:
            pass
        return []

    def _save_recent_files(self):
        try:
            RECENT_FILES_PATH.parent.mkdir(parents=True, exist_ok=True)
            RECENT_FILES_PATH.write_text(json.dumps(self._recent_files))
        except Exception:
            pass

    def _add_to_recent(self, path: str):
        if path in self._recent_files:
            self._recent_files.remove(path)
        self._recent_files.insert(0, path)
        self._recent_files = self._recent_files[:MAX_RECENT]
        self._save_recent_files()
        self._update_recent_menu()

    def _clear_recent(self):
        self._recent_files = []
        self._save_recent_files()
        self._update_recent_menu()

    # =========================================================================
    # Color meanings
    # =========================================================================
    def _load_meanings(self) -> dict:
        try:
            if MEANINGS_PATH.exists():
                return json.loads(MEANINGS_PATH.read_text())
        except Exception:
            pass
        return {}

    def _save_meanings(self):
        try:
            MEANINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
            MEANINGS_PATH.write_text(
                json.dumps(self._color_meanings, indent=2, ensure_ascii=False)
            )
        except Exception:
            pass

    # =========================================================================
    # UI construction
    # =========================================================================
    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("aqua" if "aqua" in style.theme_names() else "clam")

        # ── Header ────────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=16, pady=(14, 6))

        lhs = tk.Frame(header, bg=BG)
        lhs.pack(side="left")
        tk.Label(lhs, text=APP_NAME, bg=BG, fg=TEXT,
                 font=(_FONT_HEAD, 22, "bold")).pack(anchor="w")
        tk.Label(lhs, text=APP_TAGLINE, bg=BG, fg=TEXT2,
                 font=(_FONT_UI, 12)).pack(anchor="w")

        rhs = tk.Frame(header, bg=BG)
        rhs.pack(side="right", anchor="center")
        _viewer_btn_label = "Open in Preview" if IS_MAC else "Open PDF"
        self.preview_open_btn = ttk.Button(rhs, text=_viewer_btn_label,
                                           command=self._open_in_preview,
                                           state="disabled")
        self.preview_open_btn.pack(side="right", padx=(6, 0))
        ttk.Button(rhs, text="Open PDF\u2026", command=self._open_file).pack(side="right")

        # ── Drop zone ─────────────────────────────────────────────────────────
        self.drop_frame = tk.Frame(self, bg=SURFACE2,
                                   highlightbackground=BORDER,
                                   highlightthickness=1, height=100)
        self.drop_frame.pack(fill="x", padx=16, pady=(0, 8))
        self.drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="\U0001f4c2  Drop PDFs here  \u00b7  or click to browse",
            bg=SURFACE2, fg=TEXT2,
            font=(_FONT_UI, 14), cursor=_CURSOR_HAND,
        )
        self.drop_label.pack(expand=True)

        for w in (self.drop_frame, self.drop_label):
            w.bind("<Button-1>", lambda e: self._open_file())
            w.bind("<Enter>", self._drop_hover_on)
            w.bind("<Leave>", self._drop_hover_off)

        # ── Filter bar ────────────────────────────────────────────────────────
        filter_bar = tk.Frame(self, bg=BG)
        filter_bar.pack(fill="x", padx=16, pady=(0, 2))

        tk.Label(filter_bar, text="Filter:", bg=BG, fg=TEXT2,
                 font=(_FONT_UI, 12)).pack(side="left", padx=(0, 6))

        self.color_var = tk.StringVar(value="All Colors")
        self.color_menu = ttk.Combobox(filter_bar, textvariable=self.color_var,
                                       state="readonly", width=13)
        self.color_menu["values"] = ["All Colors"]
        self.color_menu.pack(side="left", padx=(0, 5))
        self.color_menu.bind("<<ComboboxSelected>>", lambda e: self._refresh_display())

        self.type_var = tk.StringVar(value="All Types")
        self.type_menu = ttk.Combobox(filter_bar, textvariable=self.type_var,
                                      state="readonly", width=13)
        self.type_menu["values"] = ["All Types"]
        self.type_menu.pack(side="left", padx=(0, 5))
        self.type_menu.bind("<<ComboboxSelected>>", lambda e: self._refresh_display())

        self.section_var = tk.StringVar(value="All Sections")
        self.section_menu = ttk.Combobox(filter_bar, textvariable=self.section_var,
                                         state="readonly", width=16)
        self.section_menu["values"] = ["All Sections"]
        self.section_menu.pack(side="left", padx=(0, 5))
        self.section_menu.bind("<<ComboboxSelected>>", lambda e: self._refresh_display())

        ttk.Button(filter_bar, text="Reset",
                   command=self._reset_filters).pack(side="left", padx=(0, 12))

        tk.Label(filter_bar, text="Search:", bg=BG, fg=TEXT2,
                 font=(_FONT_UI, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh_display())
        ttk.Entry(filter_bar, textvariable=self.search_var, width=16).pack(
            side="left", padx=(0, 4))
        ttk.Button(filter_bar, text="\u2715", width=2,
                   command=lambda: self.search_var.set("")).pack(side="left")

        # ── Options bar ───────────────────────────────────────────────────────
        opts = tk.Frame(self, bg=BG)
        opts.pack(fill="x", padx=16, pady=(4, 6))

        ttk.Checkbutton(opts, text="Auto-Copy", variable=self._auto_copy
                        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(opts, text="With Context", variable=self._with_context
                        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(opts, text="Deep Links", variable=self._deep_links,
                        command=self._refresh_display).pack(side="left", padx=(0, 10))

        tk.Label(opts, text="Format:", bg=BG, fg=TEXT2,
                 font=(_FONT_UI, 12)).pack(side="left", padx=(0, 4))
        fmt_menu = ttk.Combobox(opts, textvariable=self._output_format,
                                values=OUTPUT_FORMATS, state="readonly", width=20)
        fmt_menu.pack(side="left", padx=(0, 10))
        fmt_menu.bind("<<ComboboxSelected>>", lambda e: self._refresh_display())

        ttk.Button(opts, text="Color Meanings\u2026",
                   command=self._show_color_settings).pack(side="left", padx=(0, 10))

        if IS_MAC:
            if PYNPUT_AVAILABLE:
                ttk.Checkbutton(opts, text="Global Hotkey (\u2318\u2325E)",
                                variable=self._hotkey_on,
                                command=self._toggle_hotkey).pack(side="left")
            else:
                tk.Label(opts, text="Global Hotkey: install pynput",
                         bg=BG, fg=TEXT3, font=(_FONT_UI, 11)).pack(side="left")

        # ── Results area ──────────────────────────────────────────────────────
        # Both frames live in the same grid cell; lift() switches between them.
        # This avoids the pack_forget/pack height-collapse bug.
        results_outer = tk.Frame(self, bg=SURFACE)
        results_outer.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        results_outer.grid_rowconfigure(0, weight=1)
        results_outer.grid_columnconfigure(0, weight=1)

        # Card view
        self.card_frame = tk.Frame(results_outer, bg=SURFACE)
        self.card_frame.grid(row=0, column=0, sticky="nsew")

        self.card_canvas = tk.Canvas(self.card_frame, bg=SURFACE, highlightthickness=0)
        card_scroll = ttk.Scrollbar(self.card_frame, orient="vertical",
                                    command=self.card_canvas.yview)
        self.card_canvas.configure(yscrollcommand=card_scroll.set)
        card_scroll.pack(side="right", fill="y")
        self.card_canvas.pack(side="left", fill="both", expand=True)

        self.inner_frame = tk.Frame(self.card_canvas, bg=SURFACE)
        self._canvas_window = self.card_canvas.create_window(
            (0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>",
            lambda e: self.card_canvas.configure(scrollregion=self.card_canvas.bbox("all")))
        self.card_canvas.bind("<Configure>",
            lambda e: self.card_canvas.itemconfig(self._canvas_window, width=e.width))
        self.card_canvas.bind("<MouseWheel>", self._scroll_canvas)

        # Raw text view
        self.text_frame = tk.Frame(results_outer, bg=SURFACE)
        self.text_frame.grid(row=0, column=0, sticky="nsew")

        self.results = tk.Text(self.text_frame, wrap="word",
                               font=(_FONT_MONO, 12), state="normal",
                               bg=SURFACE, relief="flat", padx=14, pady=14, fg=TEXT)
        self.results.bind("<Key>",       lambda e: "break")
        self.results.bind("<BackSpace>", lambda e: "break")
        self.results.bind(f"<{_MOD}-a>", lambda e: None)
        self.results.bind(f"<{_MOD}-c>", lambda e: None)
        raw_scroll = ttk.Scrollbar(self.text_frame, orient="vertical",
                                   command=self.results.yview)
        self.results.configure(yscrollcommand=raw_scroll.set)
        raw_scroll.pack(side="right", fill="y")
        self.results.pack(side="left", fill="both", expand=True)

        # Card view on top by default
        self.card_frame.lift()

        # ── Bottom bar ────────────────────────────────────────────────────────
        bottom = tk.Frame(self, bg=BG)
        bottom.pack(fill="x", padx=16, pady=(0, 14))

        self.count_label = tk.Label(bottom, text="Open a PDF to get started",
                                    bg=BG, fg=TEXT2, font=("SF Pro Text", 12))
        self.count_label.pack(side="left")

        self.copy_feedback = tk.Label(bottom, text="", bg=BG, fg=GREEN,
                                      font=("SF Pro Text", 12, "bold"))
        self.copy_feedback.pack(side="left", padx=8)

        self.copy_btn = ttk.Button(bottom, text="Copy to Clipboard",
                                   command=self._copy_to_clipboard, state="disabled")
        self.copy_btn.pack(side="right", padx=(6, 0))

        self.save_btn = ttk.Button(bottom, text="Save As\u2026",
                                   command=self._save_as, state="disabled")
        self.save_btn.pack(side="right", padx=(6, 0))

        self.preview_btn = ttk.Button(bottom, text="Raw View",
                                      command=self._toggle_view)
        self.preview_btn.pack(side="right")

    # =========================================================================
    # Drop zone
    # =========================================================================
    def _drop_hover_on(self, _=None):
        self.drop_frame.configure(bg="#E2DDD9", highlightbackground=ACCENT)
        self.drop_label.configure(bg="#E2DDD9")

    def _drop_hover_off(self, _=None):
        self.drop_frame.configure(bg=SURFACE2, highlightbackground=BORDER)
        self.drop_label.configure(bg=SURFACE2)

    def _try_enable_dnd(self):
        try:
            import tkinterdnd2  # noqa: F401
            self.drop_frame.drop_target_register("DND_Files")       # type: ignore
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)     # type: ignore
        except Exception:
            self.drop_label.configure(text="Click Open PDF\u2026 to select a file")

    def _on_drop(self, event):
        raw = event.data.strip()
        paths = re.findall(r'\{[^}]+\}|\S+', raw)
        pdf_paths = [p.strip("{}") for p in paths if p.strip("{}").lower().endswith(".pdf")]
        if pdf_paths:
            self._load_pdfs(pdf_paths)

    # =========================================================================
    # macOS events
    # =========================================================================
    def _mac_open_doc(self, *args):
        pdf_paths = [p for p in args if p.lower().endswith(".pdf")]
        if pdf_paths:
            self._load_pdfs(pdf_paths)

    def _on_close(self):
        if self._hotkey_on.get():
            self.withdraw()
        else:
            self.destroy()

    # =========================================================================
    # File operations
    # =========================================================================
    def _open_file(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF(s)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if paths:
            self._load_pdfs(list(paths))

    def _clear_pdf(self):
        self.pdf_path       = None
        self.batch_mode     = False
        self.batch_results  = []
        self.annotations    = []
        self._metadata      = None
        self._word_counts   = None
        self.title(APP_NAME)
        self.drop_label.configure(
            text="\U0001f4c2  Drop PDFs here  \u00b7  or click to browse", fg=TEXT2)
        self.count_label.configure(text="Open a PDF to get started")
        self._set_action_buttons("disabled")
        self.color_var.set("All Colors")
        self.type_var.set("All Types")
        self.section_var.set("All Sections")
        self.search_var.set("")
        self._clear_cards()
        self.results.delete("1.0", "end")

    def _load_pdfs(self, paths: list[str]):
        self.batch_mode = len(paths) > 1
        self.annotations    = []
        self._metadata      = None
        self._word_counts   = None
        self.batch_results  = []

        if self.batch_mode:
            label = f"Processing {len(paths)} PDFs\u2026"
            self.title(f"{len(paths)} PDFs \u2014 {APP_NAME}")
            self.pdf_path = None
        else:
            name  = Path(paths[0]).name
            label = f"Processing {name}\u2026"
            self.title(f"{name} \u2014 {APP_NAME}")
            self.pdf_path = paths[0]

        self.drop_label.configure(text=label, fg=TEXT2)
        self.count_label.configure(text="Loading\u2026")
        self._set_action_buttons("disabled")
        self._clear_cards()

        with_ctx = self._with_context.get()

        def _worker():
            results = []
            for i, path in enumerate(paths):
                try:
                    # Open the PDF once for all three operations
                    anns, meta, wc = extract_all(path, with_context=with_ctx)
                    results.append({
                        "filename":    Path(path).name,
                        "path":        path,
                        "annotations": anns,
                        "metadata":    meta,
                        "word_counts": wc,
                    })
                except Exception as exc:
                    results.append({
                        "filename":    Path(path).name,
                        "path":        path,
                        "annotations": [],
                        "metadata":    {"filename": Path(path).name},
                        "word_counts": None,
                        "error":       str(exc),
                    })
                if self.batch_mode:
                    msg = f"Processing {i + 1}/{len(paths)}: {Path(path).name}\u2026"
                    self.after(0, lambda m=msg: self.drop_label.configure(text=m))
            self.after(0, lambda: self._on_load_done(paths, results))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_load_done(self, paths: list[str], results: list[dict]):
        self.batch_results = results

        flat = []
        for r in results:
            for a in r["annotations"]:
                entry = dict(a)
                entry["filename"] = r["filename"]   # always populate for CSV
                flat.append(entry)
        self.annotations = flat

        if self.batch_mode:
            label = f"{len(paths)} PDFs loaded \u2014 {len(flat)} annotations total"
            self.title(f"{len(paths)} PDFs \u2014 {APP_NAME}")
        else:
            name  = Path(paths[0]).name
            label = f"\U0001f4c4 {name}"
            self._metadata    = results[0]["metadata"]    if results else None
            self._word_counts = results[0]["word_counts"] if results else None
            for path in paths:
                self._add_to_recent(path)

        self.drop_label.configure(text=label, fg=TEXT)
        self.preview_open_btn.configure(
            state="normal" if not self.batch_mode else "disabled")
        self._rebuild_filter_menus()
        self.color_var.set("All Colors")
        self.type_var.set("All Types")
        self.section_var.set("All Sections")
        self.search_var.set("")
        self._refresh_display()

        if self._auto_copy.get() and self.annotations:
            self._copy_to_clipboard()

    # =========================================================================
    # Filters
    # =========================================================================
    def _rebuild_filter_menus(self):
        colors   = sorted({a["color"] for a in self.annotations})
        types    = sorted({a["type"]  for a in self.annotations})
        sections = sorted({a.get("section", "") for a in self.annotations
                           if a.get("section", "")})
        self.color_menu["values"]   = ["All Colors"]   + colors
        self.type_menu["values"]    = ["All Types"]    + types
        self.section_menu["values"] = ["All Sections"] + sections

    def _reset_filters(self):
        self.color_var.set("All Colors")
        self.type_var.set("All Types")
        self.section_var.set("All Sections")
        self.search_var.set("")
        self._refresh_display()

    def _filtered(self) -> list[dict]:
        result = self.annotations

        c = self.color_var.get()
        if c != "All Colors":
            result = [a for a in result if a.get("color") == c]

        t = self.type_var.get()
        if t != "All Types":
            result = [a for a in result if a["type"] == t]

        s = self.section_var.get()
        if s != "All Sections":
            result = [a for a in result if a.get("section", "") == s]

        q = self.search_var.get().strip().lower()
        if q:
            def _matches(a: dict) -> bool:
                if q in a["text"].lower():
                    return True
                if q in a.get("section", "").lower():
                    return True
                # also search the colour meaning the user assigned
                meaning = self._color_meanings.get(a.get("color", ""), "")
                if q in meaning.lower():
                    return True
                return False
            result = [a for a in result if _matches(a)]

        return result

    # =========================================================================
    # Display
    # =========================================================================
    def _refresh_display(self):
        filtered = self._filtered()

        if self._raw_mode:
            self._rebuild_raw_text(filtered)
        else:
            self._rebuild_cards(filtered)

        total = len(self.annotations)
        shown = len(filtered)
        has   = bool(filtered)

        if not self.annotations:
            self.count_label.configure(text="No annotations found in this PDF")
        elif shown == total:
            self.count_label.configure(text=f"Total: {total} annotations")
        else:
            self.count_label.configure(text=f"{shown} / {total} annotations shown")

        self._set_action_buttons("normal" if has else "disabled")

    # ── Card view ─────────────────────────────────────────────────────────────
    def _clear_cards(self):
        for child in self.inner_frame.winfo_children():
            child.destroy()

    def _rebuild_cards(self, filtered: list[dict]):
        self._clear_cards()

        if not filtered:
            msg = ("No annotations found." if not self.annotations
                   else "No annotations match the current filters.")
            tk.Label(self.inner_frame, text=f"\n\n{msg}",
                     fg=TEXT3, bg=SURFACE, font=("SF Pro Text", 14)).pack(expand=True)
            return

        if self.batch_mode:
            groups: dict[str, list[dict]] = {}
            order: list[str] = []
            for a in filtered:
                fn = a.get("filename", "Unknown")
                if fn not in groups:
                    groups[fn] = []
                    order.append(fn)
                groups[fn].append(a)
            for fn in order:
                self._build_section_header(self.inner_frame, f"\U0001f4c4 {fn}",
                                           len(groups[fn]))
                for a in sorted(groups[fn], key=lambda x: x["page"]):
                    self._build_card(self.inner_frame, a)
        else:
            buckets: dict[str, list[dict]] = {}
            b_order: list[tuple[str, str]] = []
            for a in filtered:
                if a["type"] == "Highlight":
                    key     = a["color"]
                    display = f"{COLOR_EMOJI.get(a['color'], '')} {a['color']} Highlights"
                else:
                    key     = a["type"]
                    display = f"{TYPE_EMOJI.get(a['type'], '')} {a['type']}s"
                if key not in buckets:
                    buckets[key] = []
                    b_order.append((key, display))
                buckets[key].append(a)
            for key, display in b_order:
                items = sorted(buckets[key], key=lambda x: x["page"])
                self._build_section_header(self.inner_frame, f"{display} ({len(items)})")
                for a in items:
                    self._build_card(self.inner_frame, a)

        self._bind_scroll(self.inner_frame)
        self.card_canvas.yview_moveto(0)

    def _build_section_header(self, parent, text: str, count: int | None = None):
        label = (f"{text} ({count})"
                 if count is not None and str(count) not in text else text)
        hdr = tk.Frame(parent, bg=BG)
        hdr.pack(fill="x", padx=0, pady=(10, 2))

        # Indigo left accent bar
        tk.Frame(hdr, bg=ACCENT, width=3).pack(side="left", fill="y")

        tk.Label(hdr, text=label, font=("SF Pro Text", 12, "bold"),
                 fg=TEXT, bg=BG, padx=10, pady=5,
                 anchor="w").pack(side="left", fill="x")

    def _build_card(self, parent: tk.Frame, annotation: dict):
        color    = annotation.get("color", "Unknown")
        ann_type = annotation["type"]
        page_num = annotation["page"]
        section  = annotation.get("section", "")
        meaning  = self._color_meanings.get(color, "")
        hex_c    = COLOR_HEX.get(color, "#C7C7CC")

        # Outer wrap with 1px border
        outer = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
        outer.pack(fill="x", padx=10, pady=3)

        # Row: coloured left bar + card body
        row = tk.Frame(outer, bg=SURFACE)
        row.pack(fill="x")

        bar = tk.Frame(row, bg=hex_c, width=4)
        bar.pack(side="left", fill="y")
        bar.pack_propagate(False)

        card = tk.Frame(row, bg=SURFACE)
        card.pack(side="left", fill="both", expand=True)

        # ── Header ────────────────────────────────────────────────────────────
        header_bg = "#F8F8F6"
        header = tk.Frame(card, bg=header_bg, padx=8, pady=5)
        header.pack(fill="x")

        header_lhs = tk.Frame(header, bg=header_bg)
        header_lhs.pack(side="left")

        emoji = (COLOR_EMOJI.get(color, "") if ann_type == "Highlight"
                 else TYPE_EMOJI.get(ann_type, ""))
        color_lbl = f"{emoji} {color}" if ann_type == "Highlight" else f"{emoji} {ann_type}"

        tk.Label(header_lhs, text=color_lbl, bg=header_bg, fg=hex_c,
                 font=("SF Pro Text", 11, "bold")).pack(side="left")

        if section:
            tk.Label(header_lhs, text=f"  \u00b7  {section}",
                     bg=header_bg, fg=TEXT3,
                     font=("SF Pro Text", 10, "italic")).pack(side="left")

        if meaning:
            tk.Label(header_lhs, text=f"  \u2014  {meaning}",
                     bg=header_bg, fg=TEXT2,
                     font=("SF Pro Text", 10, "italic")).pack(side="left")

        header_rhs = tk.Frame(header, bg=header_bg)
        header_rhs.pack(side="right")

        tk.Label(header_rhs, text=f"p.\u2009{page_num}",
                 bg="#EDECEA", fg=TEXT2,
                 font=("SF Pro Text", 10), padx=6, pady=1).pack(side="left", padx=(0, 8))

        tk.Button(header_rhs, text="Edit",
                  font=("SF Pro Text", 10), fg=ACCENT, bg=header_bg,
                  bd=0, cursor="pointinghand", relief="flat",
                  command=lambda a=annotation: self._edit_annotation(a)
                  ).pack(side="left", padx=(0, 4))

        tk.Button(header_rhs, text="\u00d7",
                  font=("SF Pro Text", 13), fg="#EF4444", bg=header_bg,
                  bd=0, cursor="pointinghand", relief="flat",
                  command=lambda a=annotation: self._delete_annotation(a)
                  ).pack(side="left")

        # ── Body ──────────────────────────────────────────────────────────────
        body = tk.Label(card, text=annotation["text"],
                        font=("SF Pro Text", 13), fg=TEXT,
                        bg=SURFACE, wraplength=640, justify="left",
                        anchor="nw", padx=12, pady=8)
        body.pack(fill="x", anchor="w")

        if annotation.get("context"):
            ctx_bg = "#FAFAF8"
            ctx_wrap = tk.Frame(card, bg=ctx_bg, padx=12, pady=6)
            ctx_wrap.pack(fill="x")
            tk.Label(ctx_wrap, text="Context", bg=ctx_bg, fg=TEXT3,
                     font=("SF Pro Text", 10, "bold")).pack(anchor="w")
            tk.Label(ctx_wrap, text=annotation["context"],
                     bg=ctx_bg, fg=TEXT2,
                     font=("SF Pro Text", 11, "italic"),
                     wraplength=640, justify="left", anchor="w").pack(anchor="w")

        # ── Hover / click ─────────────────────────────────────────────────────
        clickables = [card, header, header_lhs, body]

        def _hover_on(_, o=outer, hc=hex_c):
            o.configure(bg=hc)

        def _hover_off(_, o=outer):
            o.configure(bg=BORDER)

        def _click(_, p=page_num):
            self._open_at_page(p)

        for w in clickables:
            w.bind("<Button-1>", _click)
            w.bind("<Enter>",    _hover_on)
            w.bind("<Leave>",    _hover_off)

    def _bind_scroll(self, widget: tk.Widget):
        widget.bind("<MouseWheel>", self._scroll_canvas)
        for child in widget.winfo_children():
            self._bind_scroll(child)

    def _scroll_canvas(self, event):
        self.card_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── Raw view ──────────────────────────────────────────────────────────────
    def _rebuild_raw_text(self, filtered: list[dict]):
        self.results.delete("1.0", "end")
        content = self._get_export_content(filtered)
        self.results.insert("end", content if content else "(nothing to show)")

    # ── Toggle ────────────────────────────────────────────────────────────────
    def _toggle_view(self):
        self._raw_mode = not self._raw_mode
        if self._raw_mode:
            self.text_frame.lift()           # bring text frame to front
            self.preview_btn.configure(text="Card View")
        else:
            self.card_frame.lift()           # bring card frame to front
            self.preview_btn.configure(text="Raw View")
        self._refresh_display()

    def _show_empty(self, msg: str):
        self._clear_cards()
        tk.Label(self.inner_frame, text=f"\n\n{msg}",
                 fg=TEXT3, bg=SURFACE, font=("SF Pro Text", 14)).pack(expand=True)

    def _show_error(self, msg: str):
        self._clear_cards()
        tk.Label(self.inner_frame, text=f"\n\nError: {msg}",
                 fg="#EF4444", bg=SURFACE, font=("SF Pro Text", 13),
                 wraplength=580, justify="center").pack(expand=True)
        self.count_label.configure(text="Error loading PDF")
        self._set_action_buttons("disabled")

    # =========================================================================
    # Color Meanings dialog
    # =========================================================================
    def _show_color_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Color Meanings")
        dlg.geometry("400x460")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        preset_frame = ttk.Frame(dlg)
        preset_frame.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Label(preset_frame, text="Preset:").pack(side="left", padx=(0, 6))
        preset_var = tk.StringVar(value="")
        preset_menu = ttk.Combobox(preset_frame, textvariable=preset_var,
                                   values=list(MEANING_PRESETS.keys()),
                                   state="readonly", width=20)
        preset_menu.pack(side="left", padx=(0, 6))

        colors_frame = ttk.Frame(dlg)
        colors_frame.pack(fill="both", expand=True, padx=12, pady=4)
        color_order = ["Yellow", "Green", "Blue", "Orange", "Red", "Purple", "Pink"]
        entries: dict[str, tk.StringVar] = {}

        for color in color_order:
            row = ttk.Frame(colors_frame)
            row.pack(fill="x", pady=3)
            swatch = tk.Canvas(row, width=18, height=18, highlightthickness=0,
                               bg=dlg.cget("bg"))
            swatch.create_oval(2, 2, 16, 16,
                               fill=COLOR_HEX.get(color, "#8E8E93"), outline="")
            swatch.pack(side="left", padx=(0, 6))
            ttk.Label(row, text=f"{color}:", width=8).pack(side="left")
            var = tk.StringVar(value=self._color_meanings.get(color, ""))
            ttk.Entry(row, textvariable=var, width=26).pack(side="left", fill="x", expand=True)
            entries[color] = var

        def _apply_preset():
            preset = preset_var.get()
            if preset and preset in MEANING_PRESETS:
                for color, meaning in MEANING_PRESETS[preset].items():
                    if color in entries:
                        entries[color].set(meaning)

        ttk.Button(preset_frame, text="Apply", command=_apply_preset).pack(side="left")

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", padx=12, pady=(6, 12))

        def _clear_all():
            for var in entries.values():
                var.set("")

        def _save():
            for color, var in entries.items():
                val = var.get().strip()
                if val:
                    self._color_meanings[color] = val
                elif color in self._color_meanings:
                    del self._color_meanings[color]
            self._save_meanings()
            self._refresh_display()
            dlg.destroy()

        ttk.Button(btn_frame, text="Clear All", command=_clear_all).pack(side="left")
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="right", padx=(4, 0))
        ttk.Button(btn_frame, text="Save", command=_save).pack(side="right")
        dlg.bind("<Escape>",        lambda e: dlg.destroy())
        dlg.bind(f"<{_MOD}-Return>", lambda e: _save())

    # =========================================================================
    # Annotation management
    # =========================================================================
    def _delete_annotation(self, annotation: dict):
        if annotation in self.annotations:
            self.annotations.remove(annotation)
            for r in self.batch_results:
                if annotation in r["annotations"]:
                    r["annotations"].remove(annotation)
                    break
        self._refresh_display()

    def _edit_annotation(self, annotation: dict):
        modal = tk.Toplevel(self)
        modal.title(f"Edit \u2014 Page {annotation['page']}")
        modal.geometry("520x280")
        modal.resizable(True, True)
        modal.transient(self)
        modal.grab_set()

        ttk.Label(modal,
                  text=f"Page {annotation['page']} \u2014 {annotation['type']}",
                  font=(_FONT_UI, 13, "bold")).pack(padx=12, pady=(12, 4), anchor="w")

        txt = tk.Text(modal, wrap="word", font=(_FONT_UI, 13),
                      padx=8, pady=8, height=7, relief="flat",
                      highlightbackground="#C7C7CC", highlightthickness=1)
        txt.insert("1.0", annotation["text"])
        txt.pack(fill="both", expand=True, padx=12, pady=4)
        txt.focus_set()

        btns = ttk.Frame(modal)
        btns.pack(fill="x", padx=12, pady=(4, 12))

        def _save():
            new = txt.get("1.0", "end-1c").strip()
            if new:
                annotation["text"] = new
                self._refresh_display()
            modal.destroy()

        ttk.Button(btns, text="Cancel", command=modal.destroy).pack(side="right", padx=(4, 0))
        ttk.Button(btns, text="Save",   command=_save).pack(side="right")
        modal.bind("<Escape>",        lambda e: modal.destroy())
        modal.bind(f"<{_MOD}-Return>", lambda e: _save())

    # =========================================================================
    # Open in viewer
    # =========================================================================
    def _open_in_preview(self):
        if not self.pdf_path:
            return
        if IS_MAC:
            subprocess.Popen(["open", self.pdf_path],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            os.startfile(self.pdf_path)

    def _reveal_in_finder(self):
        if not self.pdf_path:
            return
        if IS_MAC:
            subprocess.Popen(["open", "-R", self.pdf_path],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        elif IS_WIN:
            subprocess.Popen(["explorer", "/select,", self.pdf_path],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def _open_at_page(self, page_num: int):
        path = self.pdf_path
        if not path:
            return
        if IS_MAC:
            safe_path = path.replace("\\", "\\\\").replace('"', '\\"')
            script = (
                'tell application "Preview"\n'
                f'    set thePath to "{safe_path}"\n'
                '    activate\n'
                '    open (POSIX file thePath)\n'
                '    delay 0.8\n'
                '    tell front window\n'
                f'        set current page to page {page_num} of document 1\n'
                '    end tell\n'
                'end tell\n'
            )
            subprocess.Popen(["osascript", "-e", script],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            # Windows/Linux: open in default PDF viewer (page jump not supported)
            os.startfile(path)

    # =========================================================================
    # Export
    # =========================================================================
    def _get_export_content(self, filtered: list[dict]) -> str:
        fmt  = self._output_format.get()
        deep = self._deep_links.get()
        path = self.pdf_path
        meta = self._metadata
        wc   = self._word_counts
        cm   = self._color_meanings

        if fmt == "CSV":
            return format_csv(filtered, cm)
        if fmt == "Markdown Table":
            return format_markdown_table(filtered, cm)
        if fmt == "HTML Report":
            return format_html(filtered,
                               meta if not self.batch_mode else None,
                               wc   if not self.batch_mode else None,
                               cm)

        if self.batch_mode:
            filtered_ids = set(id(a) for a in filtered)
            batch = []
            for r in self.batch_results:
                anns = [a for a in r["annotations"] if id(a) in filtered_ids]
                batch.append({**r, "annotations": anns})
            with_meta = fmt in ("Markdown + Metadata", "Obsidian Callouts")
            return format_batch_markdown(batch, with_metadata=with_meta, deep_links=deep)

        if fmt == "Markdown + Metadata" and meta:
            return format_markdown_with_metadata(filtered, meta, wc, path, deep)
        if fmt == "Obsidian Callouts":
            return format_obsidian(filtered, meta, path, deep)
        return format_markdown(filtered, path, deep)

    def _set_action_buttons(self, state: str):
        self.copy_btn.configure(state=state)
        self.save_btn.configure(state=state)

    # =========================================================================
    # Save / copy
    # =========================================================================
    def _save_as(self):
        filtered = self._filtered()
        if not filtered:
            return

        fmt  = self._output_format.get()
        stem = Path(self.pdf_path).stem if self.pdf_path else "annotations"

        if fmt == "CSV":
            default_ext = ".csv"
            filetypes   = [("CSV", "*.csv"), ("All Files", "*.*")]
        elif fmt == "HTML Report":
            default_ext = ".html"
            filetypes   = [("HTML", "*.html"), ("All Files", "*.*")]
        else:
            default_ext = ".md"
            filetypes   = [("Markdown", "*.md"), ("Plain Text", "*.txt"),
                           ("JSON", "*.json")]

        save_path = filedialog.asksaveasfilename(
            title="Save Annotations",
            initialfile=stem,
            defaultextension=default_ext,
            filetypes=filetypes,
        )
        if not save_path:
            return

        try:
            if save_path.endswith(".json"):
                clean = [{k: v for k, v in a.items() if k != "color_rgb"}
                         for a in filtered]
                content = json.dumps(clean, indent=2, ensure_ascii=False)
            elif save_path.endswith(".csv"):
                content = format_csv(filtered, self._color_meanings)
            elif save_path.endswith(".html"):
                content = format_html(filtered,
                                      self._metadata    if not self.batch_mode else None,
                                      self._word_counts if not self.batch_mode else None,
                                      self._color_meanings)
            else:
                content = self._get_export_content(filtered)

            Path(save_path).write_text(content, encoding="utf-8")
            self.copy_feedback.configure(text=f"Saved to {Path(save_path).name}")
            self.after(3000, lambda: self.copy_feedback.configure(text=""))
        except Exception as exc:
            messagebox.showerror("Save Failed", str(exc))

    def _copy_to_clipboard(self):
        filtered = self._filtered()
        if not filtered:
            return
        content = self._get_export_content(filtered)
        self.clipboard_clear()
        self.clipboard_append(content)
        self.copy_feedback.configure(text="Copied!")
        self.after(2000, lambda: self.copy_feedback.configure(text=""))

    # =========================================================================
    # Global hotkey
    # =========================================================================
    def _toggle_hotkey(self):
        self._setup_hotkey(self._hotkey_on.get())

    def _setup_hotkey(self, enable: bool):
        if not PYNPUT_AVAILABLE:
            return
        if not enable:
            if self._hotkey_listener:
                self._hotkey_listener.stop()
                self._hotkey_listener = None
            return

        COMBO = {_kb.Key.cmd, _kb.Key.alt, _kb.KeyCode.from_char("e")}
        current: set = set()

        def _press(key):
            try:
                current.add(key)
            except Exception:
                pass
            if all(k in current for k in COMBO):
                self.after(0, self._run_hotkey_extraction)

        def _release(key):
            current.discard(key)

        self._hotkey_listener = _kb.Listener(on_press=_press, on_release=_release)
        self._hotkey_listener.daemon = True
        self._hotkey_listener.start()

    def _run_hotkey_extraction(self):
        if not IS_MAC:
            return
        script = """
tell application "Finder"
    set sel to selection as alias list
    if sel is {} then return ""
    set f to item 1 of sel
    if name extension of f is "pdf" then
        return POSIX path of f
    end if
    return ""
end tell
"""
        result = subprocess.run(["osascript", "-e", script],
                                capture_output=True, text=True)
        pdf_path = result.stdout.strip()
        if not pdf_path:
            self._notify(APP_NAME, "No PDF selected in Finder.")
            return

        with_ctx = self._with_context.get()

        def _worker():
            try:
                from extract import extract_annotations
                anns = extract_annotations(pdf_path, with_context=with_ctx)
                md   = format_markdown(anns)
                self.after(0, lambda: self._finish_hotkey(pdf_path, anns, md))
            except Exception as exc:
                self.after(0, lambda: self._notify("Error", str(exc)))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_hotkey(self, pdf_path: str, anns: list[dict], md: str):
        self.clipboard_clear()
        self.clipboard_append(md)
        n = len(anns)
        self._notify(APP_NAME,
                     f"Extracted {n} annotation{'s' if n != 1 else ''}"
                     f" from {Path(pdf_path).name}")

    def _notify(self, title: str, message: str):
        if IS_MAC:
            t = title.replace('"', '\\"')
            m = message.replace('"', '\\"')
            subprocess.Popen(
                ["osascript", "-e",
                 f'display notification "{m}" with title "{t}"'],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        # On Windows the status bar provides sufficient feedback; no-op here.


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
