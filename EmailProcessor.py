import Rectangulator
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from math import ceil
from concurrent.futures import ThreadPoolExecutor
from contextlib import nullcontext
import threading
import hashlib
import shutil
import traceback
import win32print
import win32gui
import win32api
import socket
import json
import ssl
import imaplib
import smtplib
import sqlite3
import email
import queue
import time
import fitz
import sys
import os
import re
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

if getattr(sys, 'frozen', False):
    # If running as a compiled .exe, use the folder containing the .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # If running as a normal Python script, use the folder containing the script
    BASE_DIR = os.path.dirname(__file__)

with open(os.path.join(BASE_DIR, "config.json")) as f:
    config = json.load(f)


# Add new settings without breaking existing config files. Missing values are
# persisted immediately so they also appear in config.json for external editing.
_template_parent = os.path.dirname(str(config.get("TEMPLATE_FOLDER", "") or "")) or BASE_DIR
_test_template_parent = os.path.dirname(str(config.get("TEST_TEMPLATE_FOLDER", "") or "")) or BASE_DIR
_invoice_parent = os.path.dirname(str(config.get("INVOICE_FOLDER", "") or "")) or BASE_DIR
_test_invoice_parent = os.path.dirname(str(config.get("TEST_INVOICE_FOLDER", "") or "")) or BASE_DIR
_CONFIG_DEFAULTS = {
    "OCR_TEMPLATE_FOLDER": os.path.join(_template_parent, "OCR_Templates"),
    "STATEMENT_TEMPLATE_FOLDER": os.path.join(_template_parent, "Statement_Templates"),
    "TEST_OCR_TEMPLATE_FOLDER": os.path.join(_test_template_parent, "OCR_Templates"),
    "TEST_STATEMENT_TEMPLATE_FOLDER": os.path.join(_test_template_parent, "Statement_Templates"),
    "TEST_STATEMENT_FOLDER": os.path.join(_test_invoice_parent, "Statements"),
    "OCR_FUZZY_THRESHOLD": 0.72,
    "OCR_DPI": 250,
    "OCR_LANGUAGE": "eng",
    "TESSDATA_PREFIX": "",
    "PRINTER_NAME": "",                 # blank = Windows default printer
    "EMAIL_WORKERS": 3,
    "NO_PDF_LABEL": "Not_Invoices",
    "MIN_EMBEDDED_TEXT_CHARS": 40,
    "POSTFIX_VENDORS": "",
}
_config_changed = False
for _key, _value in _CONFIG_DEFAULTS.items():
    if _key not in config:
        config[_key] = _value
        _config_changed = True
if _config_changed:
    with open(os.path.join(BASE_DIR, "config.json"), "w", encoding="utf-8") as _f:
        json.dump(config, _f, indent=2)


class RedirectText:

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.after(0, self.text_widget.insert, tk.END, message)
        self.text_widget.after(0, self.text_widget.see, tk.END)  # auto-scroll to the end

    def flush(self):
        pass


class ToolTip:
    """Small dependency-free tooltip used to explain controls on hover."""

    def __init__(self, widget, text, delay=450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._job = None
        self._window = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        try:
            self._job = self.widget.after(self.delay, self._show)
        except tk.TclError:
            pass

    def _cancel(self):
        if self._job is not None:
            try:
                self.widget.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None

    def _show(self):
        self._job = None
        if self._window is not None or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
            tip = tk.Toplevel(self.widget)
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{x}+{y}")
            tk.Label(
                tip, text=self.text, justify=tk.LEFT, wraplength=320,
                bg="#253238", fg="#F5F7F8", padx=9, pady=6,
                font=("Segoe UI", 9), relief=tk.FLAT,
            ).pack()
            self._window = tip
        except tk.TclError:
            self._window = None

    def _hide(self, _event=None):
        self._cancel()
        if self._window is not None:
            try:
                self._window.destroy()
            except tk.TclError:
                pass
            self._window = None


class EmailProcessor:

    # CONSTANTS
    CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
    ICON_PATH = os.path.join(BASE_DIR, "Hotpot.ico")
    TEMPLATE_FOLDER = config["TEMPLATE_FOLDER"]
    OCR_TEMPLATE_FOLDER = config["OCR_TEMPLATE_FOLDER"]
    STATEMENT_TEMPLATE_FOLDER = config["STATEMENT_TEMPLATE_FOLDER"]
    STATEMENT_FOLDER = config["STATEMENT_FOLDER"]
    INVOICE_FOLDER = config["INVOICE_FOLDER"]
    EMAIL_ENDING = config["EMAIL_ENDING"]
    ARCHIVE_DB = os.path.join(BASE_DIR, "archive.db")

    def __init__(self, username, password):
        try:
            # VARIABLES
            socket.setdefaulttimeout(100)  # set default socket timeout
            self.username = username
            self.password = password
            self.window_closed = None
            self.processor_thread = None
            self.processor_running = False
            self.pause_event = threading.Event()  # used for cycles
            self.connected = False
            self.logging_out = False
            self.TESTING = False
            self.AWAY_MODE = False
            self.current_emails = set()  # set of emails that are currently being processed
            self.current_emails_lock = threading.Lock()  # lock for current_emails
            self.remaining_pdfs = {}  # set of pdfs that are still being processed per uid
            self.valid_invoice_flags = {} # dictionary to track if an invoice is valid per uid
            self.state_lock = threading.RLock()
            self.email_executor = None

            # GUI -----------------------------------------------------------------
            # Pewter deliberately uses a restrained steel / blue-gray palette.  The
            # goal is to make state obvious without turning the application into a
            # wall of saturated status colors.
            self.palette = {
                "bg": "#E9EEF1",
                "surface": "#FFFFFF",
                "surface_alt": "#F5F7F8",
                "border": "#D4DDE1",
                "text": "#26343A",
                "muted": "#6B7A81",
                "pewter": "#77878E",
                "pewter_dark": "#4F6067",
                "accent": "#3E7180",
                "accent_hover": "#335E6A",
                "success": "#3E7F66",
                "success_soft": "#E4F2EB",
                "warning": "#B67C28",
                "warning_soft": "#FFF1D7",
                "danger": "#B55353",
                "danger_soft": "#FBE7E7",
                "info_soft": "#E5F0F4",
                "purple": "#766A91",
                "console": "#1F282C",
                "console_text": "#DCE4E7",
            }

            self.root = tk.Tk()
            try:
                self.root.iconbitmap(self.ICON_PATH)
            except (tk.TclError, OSError):
                pass
            self.root.configure(bg=self.palette["bg"])

            # Size relative to the actual screen, then start maximized so the PDF
            # review pane gets as much room as possible.
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            self.root.geometry(f"{int(screen_w * 0.95)}x{int(screen_h * 0.90)}+0+0")
            self.root.minsize(1180, 700)
            try:
                self.root.state("zoomed")
            except tk.TclError:
                pass
            self.root.title(f"Pewter — {username.upper()}")

            style = ttk.Style(self.root)
            style.theme_use("clam")
            style.configure("TNotebook", background=self.palette["bg"], borderwidth=0)
            style.configure(
                "TNotebook.Tab", background=self.palette["surface_alt"],
                foreground=self.palette["muted"], padding=(16, 8),
                font=("Segoe UI Semibold", 9), borderwidth=0,
            )
            style.map(
                "TNotebook.Tab",
                background=[("selected", self.palette["surface"])],
                foreground=[("selected", self.palette["accent"])],
            )
            style.configure(
                "Treeview", background=self.palette["surface"],
                fieldbackground=self.palette["surface"], foreground=self.palette["text"],
                rowheight=28, borderwidth=0, font=("Segoe UI", 9),
            )
            style.map(
                "Treeview",
                background=[("selected", self.palette["info_soft"])],
                foreground=[("selected", self.palette["text"])],
            )
            style.configure(
                "Treeview.Heading", background=self.palette["surface_alt"],
                foreground=self.palette["pewter_dark"], relief="flat",
                font=("Segoe UI Semibold", 9), padding=(6, 7),
            )
            style.map("Treeview.Heading", background=[("active", self.palette["info_soft"])])
            style.configure(
                "Vertical.TScrollbar", background=self.palette["surface_alt"],
                troughcolor=self.palette["bg"], bordercolor=self.palette["bg"],
                arrowcolor=self.palette["pewter_dark"],
            )

            def make_action_button(parent, text, command, tone="neutral", state=tk.NORMAL,
                                   width=None, compact=False):
                tones = {
                    "neutral": (self.palette["surface_alt"], self.palette["text"], "#E8EDF0"),
                    "primary": (self.palette["accent"], "#FFFFFF", self.palette["accent_hover"]),
                    "danger": (self.palette["danger_soft"], self.palette["danger"], "#F5D5D5"),
                    "warning": (self.palette["warning_soft"], "#8A5A17", "#F7E3BD"),
                    "success": (self.palette["success_soft"], self.palette["success"], "#D3E9DE"),
                }
                bg, fg, active = tones[tone]
                button_options = dict(
                    text=text, command=command, state=state,
                    bg=bg, fg=fg, activebackground=active, activeforeground=fg,
                    disabledforeground="#9BA7AC", relief=tk.FLAT, bd=0,
                    highlightthickness=0, padx=8 if compact else 11,
                    pady=4 if compact else 6, cursor="hand2",
                    font=("Segoe UI Semibold", 8 if compact else 9),
                )
                if width is not None:
                    button_options["width"] = width
                button = tk.Button(parent, **button_options)
                button._pewter_base_bg = bg
                button._pewter_base_fg = fg
                return button

            def add_tip(widget, text):
                ToolTip(widget, text)
                return widget

            # Notebook and application header -------------------------------------
            notebook = ttk.Notebook(self.root)
            notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            program_tab = tk.Frame(notebook, bg=self.palette["bg"])
            notebook.add(program_tab, text="  Pewter  ")

            # Keep connection/mode state, but remove the decorative PEWTER title
            # bar so the working dashboard begins immediately below the tab strip.
            self.connection_status_var = tk.StringVar(value="● Ready")
            self.mode_status_var = tk.StringVar(value="Standard mode")

            # Alert popup is layered over the program tab when Rectangulator needs
            # a yes/no decision.
            self.alert_container = tk.Frame(
                program_tab, relief="flat", bg=self.palette["surface"],
                highlightthickness=1, highlightbackground=self.palette["border"],
            )
            self.alert_container.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            self.alert_container.lower()

            # Main split: operational dashboard on the left, invoice review on the right.
            self.main_paned = tk.PanedWindow(
                program_tab, orient=tk.HORIZONTAL, sashwidth=8,
                sashrelief="flat", opaqueresize=False, bd=0,
                bg=self.palette["border"],
            )
            self.main_paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8)
            left_frame = tk.Frame(self.main_paned, bg=self.palette["bg"])
            right_frame = tk.Frame(self.main_paned, bg=self.palette["surface"])
            self.main_paned.add(left_frame, minsize=500, stretch="never")
            self.main_paned.add(right_frame, minsize=560, stretch="always")

            # Processing controls -------------------------------------------------
            controls_card = tk.Frame(
                left_frame, bg=self.palette["surface"], highlightthickness=1,
                highlightbackground=self.palette["border"],
            )
            controls_card.pack(side=tk.TOP, fill=tk.X, padx=(0, 6), pady=(0, 7))

            controls_header = tk.Frame(controls_card, bg=self.palette["surface"])
            controls_header.pack(fill=tk.X, padx=11, pady=(9, 4))
            tk.Label(
                controls_header, text="PROCESSING", bg=self.palette["surface"],
                fg=self.palette["pewter_dark"], font=("Segoe UI Semibold", 9),
            ).pack(side=tk.LEFT)
            controls_status = tk.Frame(controls_header, bg=self.palette["surface"])
            controls_status.pack(side=tk.RIGHT)
            self.connection_status_label = tk.Label(
                controls_status, textvariable=self.connection_status_var,
                bg=self.palette["surface"], fg=self.palette["pewter"],
                font=("Segoe UI Semibold", 8), anchor="e",
            )
            self.connection_status_label.pack(anchor="e")
            self.mode_status_label = tk.Label(
                controls_status, textvariable=self.mode_status_var,
                bg=self.palette["surface"], fg=self.palette["muted"],
                font=("Segoe UI", 7), anchor="e",
            )
            self.mode_status_label.pack(anchor="e")

            primary_controls = tk.Frame(controls_card, bg=self.palette["surface"])
            primary_controls.pack(fill=tk.X, padx=10, pady=(2, 5))
            for col in range(3):
                primary_controls.grid_columnconfigure(col, weight=1)

            self.start_button = make_action_button(
                primary_controls, "▶  Start", self.main, tone="primary")
            self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
            add_tip(self.start_button, "Connect to Gmail and begin watching for invoice emails.")

            self.pause_button = make_action_button(
                primary_controls, "Ⅱ  Pause", self.pause_processing,
                state=tk.DISABLED)
            self.pause_button.grid(row=0, column=1, sticky="ew", padx=4)
            add_tip(self.pause_button, "Temporarily stop inbox polling without logging out.")

            self.logout_button = make_action_button(
                primary_controls, "■  Stop", self.logout, tone="danger",
                state=tk.DISABLED)
            self.logout_button.grid(row=0, column=2, sticky="ew", padx=(4, 0))
            add_tip(self.logout_button, "Stop processing and disconnect the main Gmail session.")

            attention_row = tk.Frame(controls_card, bg=self.palette["surface"])
            attention_row.pack(fill=tk.X, padx=10, pady=(1, 5))
            attention_row.grid_columnconfigure(0, weight=1)
            attention_row.grid_columnconfigure(1, weight=1)
            self.errors_button = make_action_button(
                attention_row, "⚠  Resolve Errors", self.resolve_errors,
                state=tk.DISABLED)
            self.errors_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
            add_tip(self.errors_button, "Move items from the Errors Gmail label back into the processing flow.")

            self.print_errors_button = make_action_button(
                attention_row, "↻  Resolve Prints", self.resolve_prints,
                state=tk.DISABLED)
            self.print_errors_button.grid(row=0, column=1, sticky="ew", padx=(4, 0))
            add_tip(self.print_errors_button, "Retry invoices that are waiting in the Need_Print Gmail label.")

            utility_row = tk.Frame(controls_card, bg=self.palette["surface"])
            utility_row.pack(fill=tk.X, padx=10, pady=(1, 9))
            for col in range(3):
                utility_row.grid_columnconfigure(col, weight=1)

            self.testing_button = make_action_button(
                utility_row, "Testing: Off", self.toggle_testing, tone="neutral", compact=True)
            self.testing_button.grid(row=0, column=0, sticky="ew", padx=(0, 3), pady=2)
            add_tip(self.testing_button, "Use test folders and suppress production side effects while enabled.")

            self.away_mode_button = make_action_button(
                utility_row, "Away: Off", self.toggle_away_mode, tone="neutral", compact=True)
            self.away_mode_button.grid(row=0, column=1, sticky="ew", padx=3, pady=2)
            add_tip(self.away_mode_button, "Away Mode prints incoming invoices without waiting for manual review.")

            self.test_rectangulator_button = make_action_button(
                utility_row, "Test Review", self.test_rectangulator, compact=True)
            self.test_rectangulator_button.grid(row=0, column=2, sticky="ew", padx=(3, 0), pady=2)
            add_tip(self.test_rectangulator_button, "Open the configured test invoice directly in the review pane.")

            self.test_inbox_button = make_action_button(
                utility_row, "Test Inbox", self.test_inbox, state=tk.DISABLED, compact=True)
            self.test_inbox_button.grid(row=1, column=0, sticky="ew", padx=(0, 3), pady=2)
            add_tip(self.test_inbox_button, "Send the saved Test_Email message through the normal inbox flow.")

            self.archive_all_button = make_action_button(
                utility_row, "Archive Done", self.archive_all, compact=True)
            self.archive_all_button.grid(row=1, column=1, sticky="ew", padx=3, pady=2)
            add_tip(self.archive_all_button, "Move every completed queue row into the local archive tab.")

            self.clear_button = make_action_button(
                utility_row, "Clear Activity", lambda: self.log_text_widget.delete("1.0", tk.END), compact=True)
            self.clear_button.grid(row=1, column=2, sticky="ew", padx=(3, 0), pady=2)
            add_tip(self.clear_button, "Clear the visible activity panel. The log file on disk is unchanged.")

            # Invoice queue --------------------------------------------------------
            inbox_card = tk.Frame(
                left_frame, bg=self.palette["surface"], highlightthickness=1,
                highlightbackground=self.palette["border"],
            )
            inbox_card.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=(0, 6), pady=(0, 7))
            inbox_header = tk.Frame(inbox_card, bg=self.palette["surface"])
            inbox_header.pack(fill=tk.X, padx=11, pady=(8, 5))
            tk.Label(
                inbox_header, text="DOCUMENT QUEUE", bg=self.palette["surface"],
                fg=self.palette["pewter_dark"], font=("Segoe UI Semibold", 9),
            ).pack(side=tk.LEFT)
            self.queue_count_var = tk.StringVar(value="0 items")
            tk.Label(
                inbox_header, textvariable=self.queue_count_var,
                bg=self.palette["surface"], fg=self.palette["muted"],
                font=("Segoe UI", 8),
            ).pack(side=tk.RIGHT)

            inbox_table = tk.Frame(inbox_card, bg=self.palette["surface"])
            inbox_table.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
            self.inbox = ttk.Treeview(
                inbox_table,
                columns=("Subject", "Date", "Invoice", "Saved", "Printed", "Errors", "Filepath"),
                show="headings", height=12,
            )
            self.inbox.column("Subject", width=155, minwidth=110, anchor="w")
            self.inbox.column("Date", width=72, minwidth=65, anchor="center")
            self.inbox.column("Invoice", width=130, minwidth=90, anchor="w")
            self.inbox.column("Saved", width=58, minwidth=52, anchor="center")
            self.inbox.column("Printed", width=58, minwidth=52, anchor="center")
            self.inbox.column("Errors", width=110, minwidth=80, anchor="w")
            self.inbox.column("Filepath", width=0, stretch=False)
            self.inbox.heading("Subject", text="Email")
            self.inbox.heading("Date", text="Received")
            self.inbox.heading("Invoice", text="Invoice / File")
            self.inbox.heading("Saved", text="Saved")
            self.inbox.heading("Printed", text="Printed")
            self.inbox.heading("Errors", text="Status")
            self.inbox.heading("Filepath", text="")
            inbox_scroll = ttk.Scrollbar(inbox_table, orient="vertical", command=self.inbox.yview)
            self.inbox.configure(yscrollcommand=inbox_scroll.set)
            self.inbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            inbox_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.inbox.bind("<Double-1>", self.remove_inbox_item)
            self.inbox.tag_configure("pending", background=self.palette["warning_soft"])
            self.inbox.tag_configure("finished", background=self.palette["success_soft"])
            self.inbox.tag_configure("error", background=self.palette["danger_soft"])
            self.inbox.tag_configure("default", background=self.palette["surface"])
            tk.Label(
                inbox_card, text="Double-click a completed row to move it to Archive.",
                bg=self.palette["surface"], fg=self.palette["muted"],
                font=("Segoe UI", 8), anchor="w",
            ).pack(fill=tk.X, padx=10, pady=(0, 7))

            # Activity log ---------------------------------------------------------
            log_card = tk.Frame(
                left_frame, bg=self.palette["surface"], highlightthickness=1,
                highlightbackground=self.palette["border"],
            )
            log_card.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=(0, 6))
            log_header = tk.Frame(log_card, bg=self.palette["surface"])
            log_header.pack(fill=tk.X, padx=11, pady=(8, 4))
            tk.Label(
                log_header, text="ACTIVITY", bg=self.palette["surface"],
                fg=self.palette["pewter_dark"], font=("Segoe UI Semibold", 9),
            ).pack(side=tk.LEFT)
            tk.Label(
                log_header, text="Newest messages appear at the bottom",
                bg=self.palette["surface"], fg=self.palette["muted"],
                font=("Segoe UI", 8),
            ).pack(side=tk.RIGHT)
            log_body = tk.Frame(log_card, bg=self.palette["surface"])
            log_body.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
            scrollbar = ttk.Scrollbar(log_body, orient="vertical")
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.log_text_widget = tk.Text(
                log_body, yscrollcommand=scrollbar.set, height=9, width=60,
                spacing1=2, spacing3=2, padx=8, pady=6, wrap=tk.WORD,
                bg="#FBFCFC", fg=self.palette["text"], insertbackground=self.palette["text"],
                selectbackground=self.palette["info_soft"], relief=tk.FLAT, bd=0,
                font=("Segoe UI", 9),
            )
            self.log_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.configure(command=self.log_text_widget.yview)

            # Invoice review pane --------------------------------------------------
            review_header = tk.Frame(right_frame, bg=self.palette["surface"], height=48)
            review_header.pack(side=tk.TOP, fill=tk.X)
            review_header.pack_propagate(False)
            review_titles = tk.Frame(review_header, bg=self.palette["surface"])
            review_titles.pack(side=tk.LEFT, padx=12, pady=7)
            tk.Label(
                review_titles, text="DOCUMENT REVIEW", bg=self.palette["surface"],
                fg=self.palette["pewter_dark"], font=("Segoe UI Semibold", 10),
            ).pack(anchor="w")

            self.figure = Figure(figsize=(9, 9.5), dpi=100, facecolor=self.palette["surface"])
            self.ax = self.figure.add_subplot(111)
            self.figure.subplots_adjust(left=0.05, right=0.95, top=0.90, bottom=0.18)
            self.ax.set_facecolor(self.palette["surface"])
            self.ax.axis("off")
            self.ax.text(
                0.5, 0.54, "PEWTER", transform=self.ax.transAxes,
                ha="center", va="center", fontsize=25, fontweight="semibold",
                color=self.palette["border"],
            )
            self.ax.text(
                0.5, 0.47, "Measuring... Measuring...", transform=self.ax.transAxes,
                ha="center", va="center", fontsize=10, color=self.palette["muted"],
            )
            self.canvas = FigureCanvasTkAgg(self.figure, master=right_frame)
            self.canvas.get_tk_widget().configure(bg=self.palette["surface"], highlightthickness=0)
            self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            self.gui_queue = queue.PriorityQueue()
            self.gui_busy = False
            self.rectangulator_handler = Rectangulator.RectangulatorHandler(self, self.figure, self.ax)

            # Archive tab ----------------------------------------------------------
            archive_tab = tk.Frame(notebook, bg=self.palette["bg"])
            notebook.add(archive_tab, text="  Archive  ")
            archive_header = tk.Frame(archive_tab, bg=self.palette["surface"])
            archive_header.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 6))
            archive_title = tk.Frame(archive_header, bg=self.palette["surface"])
            archive_title.pack(side=tk.LEFT, padx=8, pady=7)
            tk.Label(
                archive_title, text="Processed documents", bg=self.palette["surface"],
                fg=self.palette["text"], font=("Segoe UI Semibold", 13),
            ).pack(anchor="w")
            archive_controls = tk.Frame(archive_header, bg=self.palette["surface"])
            archive_controls.pack(side=tk.RIGHT, padx=8, pady=8)
            reprocess_button = make_action_button(
                archive_controls, "↻  Reprocess", self.reprocess_archive_item, tone="warning", compact=True)
            reprocess_button.pack(side=tk.LEFT, padx=3)
            open_archive_button = make_action_button(
                archive_controls, "Open File", lambda: self.open_archive_item(None), compact=True)
            open_archive_button.pack(side=tk.LEFT, padx=3)
            add_tip(reprocess_button, "Make a working copy of the selected file and process it again.")
            add_tip(open_archive_button, "Open the selected archived PDF with the Windows default viewer.")

            archive_body = tk.Frame(archive_tab, bg=self.palette["surface"])
            archive_body.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
            self.archive = ttk.Treeview(
                archive_body,
                columns=("Subject", "Date", "Invoice", "Saved", "Printed", "Errors", "Filepath"),
                show="headings", height=15,
            )
            self.archive.column("Subject", width=160, anchor="w")
            self.archive.column("Date", width=85, anchor="center")
            self.archive.column("Invoice", width=180, anchor="w")
            self.archive.column("Saved", width=65, anchor="center")
            self.archive.column("Printed", width=65, anchor="center")
            self.archive.column("Errors", width=145, anchor="w")
            self.archive.column("Filepath", width=420, anchor="w")
            self.archive.heading("Subject", text="Email")
            self.archive.heading("Date", text="Date")
            self.archive.heading("Invoice", text="Invoice")
            self.archive.heading("Saved", text="Saved")
            self.archive.heading("Printed", text="Printed")
            self.archive.heading("Errors", text="Status")
            self.archive.heading("Filepath", text="File")
            self.archive.bind("<Double-1>", self.open_archive_item)
            self.archive.bind("<Double-Button-3>", self.remove_archive_item)
            self.archive.bind("<Button-2>", self.print_archive_item)
            self.archive.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.archive_scrollbar = ttk.Scrollbar(
                archive_body, orient="vertical", command=self.archive.yview)
            self.archive_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.archive.configure(yscrollcommand=self.archive_scrollbar.set)

            # Local databases ------------------------------------------------------
            self.db = sqlite3.connect(self.ARCHIVE_DB)
            self.db.execute("""CREATE TABLE IF NOT EXISTS archive (
                     id       TEXT PRIMARY KEY,
                     subject  TEXT,
                     datestamp TEXT,
                     invoice  TEXT,
                     saved    TEXT,
                     printed  TEXT,
                     errors   TEXT,
                     filepath TEXT
                   )""")
            self.db.execute("""CREATE TABLE IF NOT EXISTS document_history (
                         sha256       TEXT PRIMARY KEY,
                         is_invoice   INTEGER NOT NULL DEFAULT 1,
                         saved        INTEGER NOT NULL DEFAULT 0,
                         printed      INTEGER NOT NULL DEFAULT 0,
                         final_name   TEXT,
                         filepath     TEXT,
                         first_seen   TEXT NOT NULL,
                         last_seen    TEXT NOT NULL
                       )""")
            self.db.commit()
            self.load_archive()

            # Console tab ----------------------------------------------------------
            console_tab = tk.Frame(notebook, bg=self.palette["console"])
            notebook.add(console_tab, text="  Console  ")
            console_header = tk.Frame(console_tab, bg=self.palette["console"])
            console_header.pack(fill=tk.X, padx=12, pady=(10, 3))
            tk.Label(
                console_header, text="Developer console", bg=self.palette["console"],
                fg="#FFFFFF", font=("Segoe UI Semibold", 11),
            ).pack(side=tk.LEFT)
            console_body = tk.Frame(console_tab, bg=self.palette["console"])
            console_body.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
            c_scrollbar = ttk.Scrollbar(console_body, orient="vertical")
            c_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            console_text = tk.Text(
                console_body, yscrollcommand=c_scrollbar.set,
                bg=self.palette["console"], fg=self.palette["console_text"],
                insertbackground="#FFFFFF", selectbackground="#38505A",
                font=("Consolas", 9), relief=tk.FLAT, padx=10, pady=8,
            )
            console_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            c_scrollbar.config(command=console_text.yview)

            # Settings tab ---------------------------------------------------------
            settings_tab = tk.Frame(notebook, bg=self.palette["bg"])
            notebook.add(settings_tab, text="  Settings  ")

            settings_top = tk.Frame(settings_tab, bg=self.palette["surface"])
            settings_top.pack(fill=tk.X, padx=10, pady=(10, 6))
            tk.Label(
                settings_top, text="Settings", bg=self.palette["surface"],
                fg=self.palette["text"], font=("Segoe UI Semibold", 13),
            ).pack(anchor="w", padx=10, pady=(8, 0))

            settings_canvas_host = tk.Frame(settings_tab, bg=self.palette["bg"])
            settings_canvas_host.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
            settings_canvas = tk.Canvas(
                settings_canvas_host, bg=self.palette["bg"], highlightthickness=0)
            settings_scroll = ttk.Scrollbar(
                settings_canvas_host, orient="vertical", command=settings_canvas.yview)
            settings_inner = tk.Frame(settings_canvas, bg=self.palette["bg"])
            settings_window = settings_canvas.create_window((0, 0), window=settings_inner, anchor="nw")
            settings_canvas.configure(yscrollcommand=settings_scroll.set)
            settings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            settings_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            settings_inner.bind(
                "<Configure>",
                lambda _e: settings_canvas.configure(scrollregion=settings_canvas.bbox("all")))
            settings_canvas.bind(
                "<Configure>",
                lambda e: settings_canvas.itemconfigure(settings_window, width=e.width))
            settings_canvas.bind(
                "<MouseWheel>",
                lambda e: settings_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

            settings = {
                "APC_USER": tk.StringVar(value=config.get("APC_USER", "")),
                "LOG_FILE": tk.StringVar(value=config.get("LOG_FILE", "")),
                "INVOICE_FOLDER": tk.StringVar(value=config.get("INVOICE_FOLDER", "")),
                "TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEMPLATE_FOLDER", "")),
                "OCR_TEMPLATE_FOLDER": tk.StringVar(value=config.get("OCR_TEMPLATE_FOLDER", "")),
                "STATEMENT_FOLDER": tk.StringVar(value=config.get("STATEMENT_FOLDER", "")),
                "STATEMENT_TEMPLATE_FOLDER": tk.StringVar(value=config.get("STATEMENT_TEMPLATE_FOLDER", "")),
                "TEST_INVOICE_FOLDER": tk.StringVar(value=config.get("TEST_INVOICE_FOLDER", "")),
                "TEST_TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEST_TEMPLATE_FOLDER", "")),
                "TEST_OCR_TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEST_OCR_TEMPLATE_FOLDER", "")),
                "TEST_STATEMENT_FOLDER": tk.StringVar(value=config.get("TEST_STATEMENT_FOLDER", "")),
                "TEST_STATEMENT_TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEST_STATEMENT_TEMPLATE_FOLDER", "")),
                "TEST_INVOICE": tk.StringVar(value=config.get("TEST_INVOICE", "")),
                "INBOX_CYCLE_TIME": tk.IntVar(value=config.get("INBOX_CYCLE_TIME", 30)),
                "RECONNECT_TIME": tk.IntVar(value=config.get("RECONNECT_TIME", 3600)),
                "RECEIVER_EMAIL": tk.StringVar(value=config.get("RECEIVER_EMAIL", "")),
                "SCANNER_EMAIL": tk.StringVar(value=config.get("SCANNER_EMAIL", "")),
                "SPLIT_VENDORS": tk.StringVar(value=config.get("SPLIT_VENDORS", "")),
                "PREFIX_VENDORS": tk.StringVar(value=config.get("PREFIX_VENDORS", "")),
                "POSTFIX_VENDORS": tk.StringVar(value=config.get("POSTFIX_VENDORS", "")),
                "PRINTER_NAME": tk.StringVar(value=config.get("PRINTER_NAME", "")),
                "OCR_FUZZY_THRESHOLD": tk.DoubleVar(value=config.get("OCR_FUZZY_THRESHOLD", 0.72)),
                "OCR_DPI": tk.IntVar(value=config.get("OCR_DPI", 250)),
                "OCR_LANGUAGE": tk.StringVar(value=config.get("OCR_LANGUAGE", "eng")),
                "TESSDATA_PREFIX": tk.StringVar(value=config.get("TESSDATA_PREFIX", "")),
                "EMAIL_WORKERS": tk.IntVar(value=config.get("EMAIL_WORKERS", 3)),
                "NO_PDF_LABEL": tk.StringVar(value=config.get("NO_PDF_LABEL", "Not_Invoices")),
                "MIN_EMBEDDED_TEXT_CHARS": tk.IntVar(value=config.get("MIN_EMBEDDED_TEXT_CHARS", 40)),
            }

            setting_meta = {
                "APC_USER": ("Fallback account user", "Used by alert-email helpers that run without the active session."),
                "LOG_FILE": ("Log file", "Plain-text activity log written alongside the on-screen activity panel."),
                "INVOICE_FOLDER": ("Invoice folder", "Where processed PDF files are saved."),
                "TEMPLATE_FOLDER": ("Text templates", "Automatic templates for PDFs that contain embedded text."),
                "OCR_TEMPLATE_FOLDER": ("OCR templates", "Human-confirmed fuzzy templates for scan-only PDFs."),
                "STATEMENT_FOLDER": ("Statements root", "Contains one destination folder per company for saved statements."),
                "STATEMENT_TEMPLATE_FOLDER": ("Statement templates", "Identifier/date templates that remember a company statement folder."),
                "TEST_INVOICE_FOLDER": ("Test invoice folder", "Invoice output location while Testing mode is enabled."),
                "TEST_TEMPLATE_FOLDER": ("Test text templates", "Text-template folder used in Testing mode."),
                "TEST_OCR_TEMPLATE_FOLDER": ("Test OCR templates", "OCR-template folder used in Testing mode."),
                "TEST_STATEMENT_FOLDER": ("Test statements root", "Statement output location while Testing mode is enabled."),
                "TEST_STATEMENT_TEMPLATE_FOLDER": ("Test statement templates", "Statement-template folder used in Testing mode."),
                "TEST_INVOICE": ("Test invoice", "PDF opened by the Test Review button."),
                "INBOX_CYCLE_TIME": ("Inbox check interval (sec)", "How often Pewter polls Gmail when no new mail is found."),
                "RECONNECT_TIME": ("Reconnect interval (sec)", "Periodic refresh interval for the long-running Gmail connection."),
                "RECEIVER_EMAIL": ("Alert recipient", "Address that receives critical Pewter error notifications."),
                "SCANNER_EMAIL": ("Scanner sender", "Messages from this address use scanner-specific behavior."),
                "SPLIT_VENDORS": ("Split vendors", "Comma-separated vendors whose multi-page PDFs should be split by page."),
                "PREFIX_VENDORS": ("Invoice prefixes", "Vendor-to-prefix rules, e.g. VendorA:VA-, VendorB:VB-."),
                "POSTFIX_VENDORS": ("Invoice postfixes", "Vendor-to-postfix rules appended after the extracted invoice number."),
                "PRINTER_NAME": ("Printer", "Leave blank to use the Windows default printer."),
                "OCR_FUZZY_THRESHOLD": ("OCR match threshold", "Minimum 0–1 similarity score before an OCR template is suggested."),
                "OCR_DPI": ("OCR DPI", "Render resolution used by Tesseract. Higher can improve small text but costs time."),
                "OCR_LANGUAGE": ("OCR language", "Tesseract language code, normally 'eng'."),
                "TESSDATA_PREFIX": ("Tesseract data folder", "Usually C:\\Program Files\\Tesseract-OCR\\tessdata on Windows."),
                "EMAIL_WORKERS": ("Email workers", "Maximum number of email-processing workers (1–10)."),
                "NO_PDF_LABEL": ("No-PDF label", "Gmail label used when a trusted email contains no PDF attachment."),
                "MIN_EMBEDDED_TEXT_CHARS": ("Minimum native text", "Below this character count Pewter treats a page as scan/OCR oriented."),
            }
            setting_groups = [
                ("Files & Templates", [
                    "INVOICE_FOLDER", "TEMPLATE_FOLDER", "OCR_TEMPLATE_FOLDER",
                    "STATEMENT_FOLDER", "STATEMENT_TEMPLATE_FOLDER", "LOG_FILE",
                ]),
                ("Email & Routing", [
                    "RECEIVER_EMAIL", "SCANNER_EMAIL", "NO_PDF_LABEL", "APC_USER",
                ]),
                ("OCR", [
                    "OCR_FUZZY_THRESHOLD", "OCR_DPI", "OCR_LANGUAGE", "TESSDATA_PREFIX",
                    "MIN_EMBEDDED_TEXT_CHARS",
                ]),
                ("Processing", [
                    "INBOX_CYCLE_TIME", "RECONNECT_TIME", "EMAIL_WORKERS", "PRINTER_NAME",
                ]),
                ("Vendor Rules", ["SPLIT_VENDORS", "PREFIX_VENDORS", "POSTFIX_VENDORS"]),
                ("Testing", [
                    "TEST_INVOICE", "TEST_INVOICE_FOLDER", "TEST_TEMPLATE_FOLDER",
                    "TEST_OCR_TEMPLATE_FOLDER", "TEST_STATEMENT_FOLDER",
                    "TEST_STATEMENT_TEMPLATE_FOLDER",
                ]),
            ]

            def browse_setting(key):
                var = settings[key]
                current = str(var.get() or "").strip()
                initial_dir = current if os.path.isdir(current) else os.path.dirname(current) or BASE_DIR
                if key.endswith("_FOLDER") or key == "INVOICE_FOLDER":
                    selected = filedialog.askdirectory(parent=self.root, initialdir=initial_dir)
                elif key == "TEST_INVOICE":
                    selected = filedialog.askopenfilename(
                        parent=self.root, initialdir=initial_dir,
                        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
                else:
                    selected = None
                if selected:
                    var.set(selected)

            for group_name, keys in setting_groups:
                group = tk.LabelFrame(
                    settings_inner, text=group_name, bg=self.palette["surface"],
                    fg=self.palette["pewter_dark"], font=("Segoe UI Semibold", 9),
                    bd=0, relief=tk.FLAT, padx=10, pady=7,
                    highlightthickness=1, highlightbackground=self.palette["border"],
                )
                group.pack(fill=tk.X, pady=(0, 8))
                group.grid_columnconfigure(1, weight=1)
                for row, key in enumerate(keys):
                    title, description = setting_meta[key]
                    label_box = tk.Frame(group, bg=self.palette["surface"])
                    label_box.grid(row=row, column=0, sticky="nw", padx=(2, 14), pady=6)
                    tk.Label(
                        label_box, text=title, bg=self.palette["surface"],
                        fg=self.palette["text"], font=("Segoe UI Semibold", 8),
                    ).pack(anchor="w")
                    tk.Label(
                        label_box, text=description, bg=self.palette["surface"],
                        fg=self.palette["muted"], font=("Segoe UI", 7),
                        wraplength=280, justify=tk.LEFT,
                    ).pack(anchor="w")
                    entry_box = tk.Frame(group, bg=self.palette["surface"])
                    entry_box.grid(row=row, column=1, sticky="ew", pady=6)
                    entry_box.grid_columnconfigure(0, weight=1)
                    entry = tk.Entry(
                        entry_box, textvariable=settings[key], bg="#FBFCFC",
                        fg=self.palette["text"], insertbackground=self.palette["text"],
                        relief=tk.FLAT, highlightthickness=1,
                        highlightbackground=self.palette["border"],
                        highlightcolor=self.palette["accent"], font=("Segoe UI", 9),
                    )
                    entry.grid(row=0, column=0, sticky="ew", ipady=5)
                    if key.endswith("_FOLDER") or key in ("INVOICE_FOLDER", "TEST_INVOICE"):
                        browse = make_action_button(
                            entry_box, "Browse", lambda k=key: browse_setting(k), compact=True)
                        browse.grid(row=0, column=1, padx=(6, 0))

            def save_settings():
                try:
                    new_values = {key: var.get() for key, var in settings.items()}
                    threshold = float(new_values.get("OCR_FUZZY_THRESHOLD", 0.72))
                    if not 0.0 <= threshold <= 1.0:
                        raise ValueError("OCR_FUZZY_THRESHOLD must be between 0.0 and 1.0")
                    if int(new_values.get("OCR_DPI", 250)) < 72:
                        raise ValueError("OCR_DPI must be at least 72")
                    if not 1 <= int(new_values.get("EMAIL_WORKERS", 3)) <= 10:
                        raise ValueError("EMAIL_WORKERS must be between 1 and 10")
                    if int(new_values.get("MIN_EMBEDDED_TEXT_CHARS", 40)) < 1:
                        raise ValueError("MIN_EMBEDDED_TEXT_CHARS must be at least 1")
                    config.update(new_values)
                    with open(self.CONFIG_PATH, "w", encoding="utf-8") as f:
                        json.dump(config, f, indent=2)
                    self.TEMPLATE_FOLDER = config.get("TEMPLATE_FOLDER", self.TEMPLATE_FOLDER)
                    self.OCR_TEMPLATE_FOLDER = config.get("OCR_TEMPLATE_FOLDER", self.OCR_TEMPLATE_FOLDER)
                    self.STATEMENT_TEMPLATE_FOLDER = config.get("STATEMENT_TEMPLATE_FOLDER", self.STATEMENT_TEMPLATE_FOLDER)
                    self.STATEMENT_FOLDER = config.get("STATEMENT_FOLDER", self.STATEMENT_FOLDER)
                    self.INVOICE_FOLDER = config.get("INVOICE_FOLDER", self.INVOICE_FOLDER)
                    for folder_key in ("TEMPLATE_FOLDER", "OCR_TEMPLATE_FOLDER", "STATEMENT_TEMPLATE_FOLDER", "STATEMENT_FOLDER", "INVOICE_FOLDER"):
                        folder = str(config.get(folder_key, "") or "").strip()
                        if folder:
                            os.makedirs(folder, exist_ok=True)
                    self.rectangulator_handler.refresh_config()
                    self.refresh_template_manager()
                    messagebox.showinfo("Settings", "Settings saved successfully.")
                except Exception as e:
                    messagebox.showerror("Settings", f"Failed to save settings: {e}")

            settings_actions = tk.Frame(settings_inner, bg=self.palette["bg"])
            settings_actions.pack(fill=tk.X, pady=(0, 4))
            save_button = make_action_button(
                settings_actions, "Save Settings", save_settings, tone="primary")
            save_button.pack(side=tk.LEFT)
            open_text_templates = make_action_button(
                settings_actions, "Open Text Templates",
                lambda: self.open_folder(config.get("TEMPLATE_FOLDER", "")), compact=True)
            open_text_templates.pack(side=tk.LEFT, padx=6)
            open_ocr_templates = make_action_button(
                settings_actions, "Open OCR Templates",
                lambda: self.open_folder(config.get("OCR_TEMPLATE_FOLDER", "")), compact=True)
            open_ocr_templates.pack(side=tk.LEFT)
            open_statement_templates = make_action_button(
                settings_actions, "Open Statement Templates",
                lambda: self.open_folder(config.get("STATEMENT_TEMPLATE_FOLDER", "")), compact=True)
            open_statement_templates.pack(side=tk.LEFT, padx=6)
            open_statements = make_action_button(
                settings_actions, "Open Statements",
                lambda: self.open_folder(config.get("STATEMENT_FOLDER", "")), compact=True)
            open_statements.pack(side=tk.LEFT)

            # Template Manager tab -------------------------------------------------
            template_tab = tk.Frame(notebook, bg=self.palette["bg"])
            notebook.add(template_tab, text="  Templates  ")
            template_header = tk.Frame(template_tab, bg=self.palette["surface"])
            template_header.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 6))
            template_title = tk.Frame(template_header, bg=self.palette["surface"])
            template_title.pack(side=tk.LEFT, padx=8, pady=7)
            tk.Label(
                template_title, text="Template Manager", bg=self.palette["surface"],
                fg=self.palette["text"], font=("Segoe UI Semibold", 13),
            ).pack(anchor="w")
            template_controls = tk.Frame(template_header, bg=self.palette["surface"])
            template_controls.pack(side=tk.RIGHT, padx=8, pady=8)
            refresh_templates_button = make_action_button(
                template_controls, "Refresh", self.refresh_template_manager, compact=True)
            refresh_templates_button.pack(side=tk.LEFT, padx=3)
            open_template_button = make_action_button(
                template_controls, "Open Selected", self.open_template_item, compact=True)
            open_template_button.pack(side=tk.LEFT, padx=3)
            delete_template_button = make_action_button(
                template_controls, "Delete Selected", self.delete_template_item,
                tone="danger", compact=True)
            delete_template_button.pack(side=tk.LEFT, padx=3)

            template_body = tk.Frame(template_tab, bg=self.palette["surface"])
            template_body.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
            self.template_tree = ttk.Treeview(
                template_body, columns=("Type", "Vendor", "File"), show="headings")
            self.template_tree.heading("Type", text="Type")
            self.template_tree.heading("Vendor", text="Vendor / Identifier")
            self.template_tree.heading("File", text="Template file")
            self.template_tree.column("Type", width=110, anchor="center")
            self.template_tree.column("Vendor", width=260, anchor="w")
            self.template_tree.column("File", width=650, anchor="w")
            self.template_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            template_scroll = ttk.Scrollbar(
                template_body, orient="vertical", command=self.template_tree.yview)
            template_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.template_tree.configure(yscrollcommand=template_scroll.set)
            self.template_tree.bind("<Double-1>", lambda event: self.open_template_item())
            self.template_items = {}
            self.refresh_template_manager()

            # Bottom status bar / crash counter -----------------------------------
            footer = tk.Frame(self.root, bg=self.palette["pewter_dark"], height=26)
            footer.pack(side=tk.BOTTOM, fill=tk.X)
            footer.pack_propagate(False)
            self.days_without_crashing = tk.StringVar()
            self.load_crash_counter()
            self.update_crash_counter_label()
            tk.Label(
                footer, textvariable=self.days_without_crashing,
                bg=self.palette["pewter_dark"], fg="#DCE5E8",
                font=("Segoe UI", 8),
            ).pack(side=tk.LEFT, padx=(10, 3))
            reset_counter_button = tk.Button(
                footer, text="↻", command=self.reset_crash_counter,
                bg=self.palette["pewter_dark"], fg="#DCE5E8",
                activebackground=self.palette["pewter"], activeforeground="#FFFFFF",
                relief=tk.FLAT, bd=0, highlightthickness=0,
                font=("Segoe UI", 8), cursor="hand2",
            )
            reset_counter_button.pack(side=tk.LEFT)
            tk.Label(
                footer, text=f"Signed in as {username.upper()}",
                bg=self.palette["pewter_dark"], fg="#AFC0C6",
                font=("Segoe UI", 8),
            ).pack(side=tk.RIGHT, padx=10)

            # Activity log styles: color is used as a cue instead of full-strength
            # neon row backgrounds, so several messages remain readable together.
            self.log_text_widget.tag_configure("red", foreground=self.palette["danger"], font=("Segoe UI Semibold", 9))
            self.log_text_widget.tag_configure("orange", foreground="#99651D")
            self.log_text_widget.tag_configure("yellow", foreground="#8A681B")
            self.log_text_widget.tag_configure("lgreen", foreground=self.palette["success"])
            self.log_text_widget.tag_configure("green", foreground=self.palette["success"], font=("Segoe UI Semibold", 9))
            self.log_text_widget.tag_configure("dgreen", foreground="#2E6E57", font=("Segoe UI Semibold", 9))
            self.log_text_widget.tag_configure("blue", foreground=self.palette["accent"])
            self.log_text_widget.tag_configure("purple", foreground=self.palette["purple"])
            self.log_text_widget.tag_configure("gray", foreground=self.palette["muted"])
            self.log_text_widget.tag_configure("no_new_emails", foreground="#8A969B")
            self.log_text_widget.tag_configure("label_error", foreground="#986717", background=self.palette["warning_soft"])
            self.log_text_widget.tag_configure("default", lmargin1=5, lmargin2=5, rmargin=5, spacing1=1, spacing3=2)

            # Redirect stdout and stderr to the dark developer console.
            sys.stdout = RedirectText(console_text)
            sys.stderr = RedirectText(console_text)

            print("""
██████╗ ███████╗██╗    ██╗████████╗███████╗██████╗
██╔══██╗██╔════╝██║    ██║╚══██╔══╝██╔════╝██╔══██╗
██████╔╝█████╗  ██║ █╗ ██║   ██║   █████╗  ██████╔╝
██╔═══╝ ██╔══╝  ██║███╗██║   ██║   ██╔══╝  ██╔══██╗
██║     ███████╗╚███╔███╔╝   ██║   ███████╗██║  ██║
╚═╝     ╚══════╝ ╚══╝╚══╝    ╚═╝   ╚══════╝╚═╝  ╚═╝
            """)

            self._refresh_ui_indicators()
            self.root.protocol("WM_DELETE_WINDOW", self.on_program_exit)  # runs exit protocol on window close
            self.root.after(1000, self.process_queue)  # starts queue processing
            self.root.mainloop()
        except Exception as e:
            print(f"-An error occurred while initializing the EmailProcessor: {str(e)}")

    def refresh_template_manager(self):
        if not hasattr(self, "template_tree"):
            return
        for item in self.template_tree.get_children():
            self.template_tree.delete(item)
        self.template_items = {}
        folders = [
            ("Text", str(config.get("TEMPLATE_FOLDER", "") or "")),
            ("OCR", str(config.get("OCR_TEMPLATE_FOLDER", "") or "")),
            ("Statement", str(config.get("STATEMENT_TEMPLATE_FOLDER", "") or "")),
        ]
        counter = 0
        for default_type, folder in folders:
            if not folder or not os.path.isdir(folder):
                continue
            for name in sorted(os.listdir(folder), key=str.casefold):
                if not name.lower().endswith((".txt", ".json")):
                    continue
                path = os.path.join(folder, name)
                vendor = os.path.splitext(name)[0]
                template_type = default_type
                if name.lower().endswith(".json"):
                    try:
                        with open(path, "r", encoding="utf-8") as f:
                            payload = json.load(f)
                        if payload.get("type") == "statement":
                            vendor = str(payload.get("company", vendor))
                            template_type = "Statement"
                        else:
                            vendor = str(payload.get("vendor", vendor))
                            template_type = "OCR" if payload.get("type") == "ocr" else "Text"
                    except Exception:
                        template_type += " (invalid JSON)"
                iid = f"template_{counter}"
                counter += 1
                self.template_items[iid] = path
                self.template_tree.insert("", "end", iid=iid, values=(template_type, vendor, path))

    def open_template_item(self):
        if not hasattr(self, "template_tree"):
            return
        selection = self.template_tree.selection()
        if not selection:
            return
        path = self.template_items.get(selection[0])
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as exc:
                messagebox.showerror("Template Manager", str(exc))

    def delete_template_item(self):
        if not hasattr(self, "template_tree"):
            return
        selection = self.template_tree.selection()
        if not selection:
            return
        path = self.template_items.get(selection[0])
        if not path or not os.path.exists(path):
            return
        if not messagebox.askyesno("Delete Template", f"Delete template?\n\n{path}"):
            return
        try:
            os.remove(path)
            if hasattr(self, "rectangulator_handler"):
                self.rectangulator_handler.invalidate_template_cache(os.path.dirname(path))
            self.refresh_template_manager()
        except Exception as exc:
            messagebox.showerror("Template Manager", str(exc))

    def _refresh_ui_indicators(self):
        """Refresh the header and mode buttons without changing processing state."""
        if not hasattr(self, "connection_status_var"):
            return

        if self.processor_running:
            if self.pause_event.is_set():
                connection_text = "● Paused"
                connection_color = self.palette["warning"]
            elif self.connected:
                connection_text = "● Connected"
                connection_color = self.palette["success"]
            else:
                connection_text = "● Connecting…"
                connection_color = self.palette["accent"]
        else:
            connection_text = "● Ready"
            connection_color = self.palette["pewter"]

        self.connection_status_var.set(connection_text)
        try:
            self.connection_status_label.config(fg=connection_color)
        except tk.TclError:
            pass

        modes = []
        if self.TESTING:
            modes.append("Testing")
        if self.AWAY_MODE:
            modes.append("Away")
        self.mode_status_var.set(" • ".join(modes) + " mode" if modes else "Standard mode")

        if hasattr(self, "testing_button"):
            if self.TESTING:
                self.testing_button.config(
                    text="Testing: On", bg=self.palette["warning_soft"],
                    fg="#8A5A17", activebackground="#F7E3BD")
            else:
                self.testing_button.config(
                    text="Testing: Off", bg=getattr(self.testing_button, "_pewter_base_bg", self.palette["surface_alt"]),
                    fg=getattr(self.testing_button, "_pewter_base_fg", self.palette["text"]),
                    activebackground="#E8EDF0")

        if hasattr(self, "away_mode_button"):
            if self.AWAY_MODE:
                self.away_mode_button.config(
                    text="Away: On", bg=self.palette["success_soft"],
                    fg=self.palette["success"], activebackground="#D3E9DE")
            else:
                self.away_mode_button.config(
                    text="Away: Off", bg=getattr(self.away_mode_button, "_pewter_base_bg", self.palette["surface_alt"]),
                    fg=getattr(self.away_mode_button, "_pewter_base_fg", self.palette["text"]),
                    activebackground="#E8EDF0")

    def _set_button_attention(self, button, active=True):
        """Highlight a recovery button when its Gmail label contains work."""
        if not button:
            return
        if active:
            button.config(
                bg=self.palette["warning_soft"], fg="#8A5A17",
                activebackground="#F7E3BD", activeforeground="#8A5A17")
        else:
            button.config(
                bg=getattr(button, "_pewter_base_bg", self.palette["surface_alt"]),
                fg=getattr(button, "_pewter_base_fg", self.palette["text"]),
                activebackground="#E8EDF0",
                activeforeground=getattr(button, "_pewter_base_fg", self.palette["text"]),
            )

    def _update_queue_count(self):
        if not hasattr(self, "queue_count_var") or not hasattr(self, "inbox"):
            return
        try:
            count = len(self.inbox.get_children())
            self.queue_count_var.set(f"{count} item" if count == 1 else f"{count} items")
        except tk.TclError:
            pass

    def _show_preview_idle(self):
        """Restore the friendly empty state after an interactive review closes."""
        if self.window_closed or self.gui_busy or not hasattr(self, "ax"):
            return
        try:
            self.ax.clear()
            self.ax.set_facecolor(self.palette["surface"])
            self.ax.axis("off")
            self.ax.text(
                0.5, 0.54, "PEWTER", transform=self.ax.transAxes,
                ha="center", va="center", fontsize=25, fontweight="semibold",
                color=self.palette["border"],
            )
            self.ax.text(
                0.5, 0.47, "No invoice waiting for review", transform=self.ax.transAxes,
                ha="center", va="center", fontsize=10, color=self.palette["muted"],
            )
            self.ax.text(
                0.5, 0.43, "New vendors, OCR documents, and exceptions will appear here.",
                transform=self.ax.transAxes, ha="center", va="center", fontsize=8,
                color=self.palette["pewter"],
            )
            self.fig.canvas.draw_idle()
        except (tk.TclError, RuntimeError):
            pass

    def main(self):
        if self.TESTING:
            self.TEMPLATE_FOLDER = config["TEST_TEMPLATE_FOLDER"]
            self.OCR_TEMPLATE_FOLDER = config.get("TEST_OCR_TEMPLATE_FOLDER", config["OCR_TEMPLATE_FOLDER"])
            self.STATEMENT_TEMPLATE_FOLDER = config.get("TEST_STATEMENT_TEMPLATE_FOLDER", config["STATEMENT_TEMPLATE_FOLDER"])
            self.STATEMENT_FOLDER = config.get("TEST_STATEMENT_FOLDER", config["STATEMENT_FOLDER"])
            self.INVOICE_FOLDER = config["TEST_INVOICE_FOLDER"]
            self.log("Testing mode enabled", tag="yellow")
        else:
            self.TEMPLATE_FOLDER = config["TEMPLATE_FOLDER"]
            self.OCR_TEMPLATE_FOLDER = config["OCR_TEMPLATE_FOLDER"]
            self.STATEMENT_TEMPLATE_FOLDER = config["STATEMENT_TEMPLATE_FOLDER"]
            self.STATEMENT_FOLDER = config["STATEMENT_FOLDER"]
            self.INVOICE_FOLDER = config["INVOICE_FOLDER"]

        for folder in (self.TEMPLATE_FOLDER, self.OCR_TEMPLATE_FOLDER, self.STATEMENT_TEMPLATE_FOLDER, self.STATEMENT_FOLDER, self.INVOICE_FOLDER):
            if folder:
                os.makedirs(folder, exist_ok=True)
        with open(config["LOG_FILE"], "a", encoding="utf-8") as file:
            file.write("\n\n")
        self.log("Connecting...", tag="dgreen")
        self.processor_running = True
        self.ui(self.button_startup)
        self.ui(self._refresh_ui_indicators)

        self.imap_lock = threading.RLock()
        try:
            self.imap = self.safe_imap(self.connect, primary=True)
        except self.IMAPUnavailable:
            self.imap = None
            self.reconnect()
        if self.imap:
            workers = max(1, min(int(config.get("EMAIL_WORKERS", 3) or 3), 10))
            if self.email_executor is not None:
                self.email_executor.shutdown(wait=False, cancel_futures=False)
            self.email_executor = ThreadPoolExecutor(max_workers=workers, thread_name_prefix="invoice-email")
            self.processor_thread = threading.Thread(target=self.search_inbox, daemon=True)
            self.processor_thread.start()

    def button_startup(self):
        # Enable and disable buttons
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(text="Ⅱ  Pause", command=self.pause_processing, state=tk.NORMAL)
        self.pause_event.clear()
        self.logout_button.config(state=tk.NORMAL)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.testing_button.config(state=tk.DISABLED)
        self.away_mode_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.NORMAL)
        self._refresh_ui_indicators()

    def button_logout(self):  
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.logout_button.config(state=tk.DISABLED)
        self.errors_button.config(state=tk.DISABLED)
        self.print_errors_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)
        self._refresh_ui_indicators()

    class IMAPUnavailable(Exception):
        pass

    def safe_imap(self, func, *args, retries=3, log_errors=True, use_lock=True, **kwargs):
        last = None
        for attempt in range(retries):
            try:
                guard = self.imap_lock if use_lock and hasattr(self, "imap_lock") else nullcontext()
                with guard:
                    return func(*args, **kwargs)
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, TimeoutError, socket.gaierror,
                    ssl.SSLError, OSError) as e:
                last = e
                if attempt < retries - 1:
                    time.sleep(min(2 ** attempt, 10))
        if log_errors:
            self.log(f"IMAP operation failed ({self.brief_error(last)})", tag="orange", console=True)
        raise self.IMAPUnavailable(str(last))

    def brief_error(self, e):  # Shortens noisy socket/SSL errors for the console
        msg = str(e)
        if "EOF occurred in violation of protocol" in msg:
            return "socket EOF"
        return msg.split("(")[0].strip()[:80]

    def connect(self, log=True, primary=True):
        user = f"{self.username}.sndex@gmail.com"
        try:
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.socket().settimeout(100)
            imap.login(user, self.password)
            if primary:
                self.connected = True
                self.ui(self._refresh_ui_indicators)
            if log:
                self.log(f"--- Connected to {self.username} --- {self.current_time} {self.current_date}", tag="dgreen")
            return imap
        except Exception as e:
            if primary:
                self.connected = False
                self.ui(self._refresh_ui_indicators)
            if log:
                self.log(f"Unable to connect to {self.username}: {e}", tag="red", send_email=True)
            raise

    def disconnect(self, imap, log=True):
        if imap is None:
            self.connected = False
            self.ui(self._refresh_ui_indicators)
            return
        try:
            imap.logout()
        except Exception as e:
            if log:
                self.log(f"An error occurred while disconnecting: {e}", tag="orange")
        finally:
            if imap is getattr(self, "imap", None):
                self.connected = False
                self.ui(self._refresh_ui_indicators)
            if log:
                self.log(f"--- Disconnected from {self.username} --- {self.current_time} {self.current_date}", tag="red")

    def search_inbox(self):
        cycle_count = 0
        while self.processor_running:
            try:
                if self.pause_event.is_set():
                    time.sleep(0.25)
                    continue
                if not self.connected or self.imap is None:
                    self.reconnect()
                    continue

                self.safe_imap(self.imap.select, "INBOX")
                _, emails = self.safe_imap(self.imap.uid, "search", None, "UNSEEN")
                uids = [uid.decode() for uid in emails[0].split() if uid]
                new_uids = []
                with self.current_emails_lock:
                    for uid in uids:
                        if uid not in self.current_emails:
                            self.current_emails.add(uid)
                            new_uids.append(uid)

                if new_uids:
                    for uid in new_uids:
                        try:
                            self.email_executor.submit(self.process_email, uid)
                        except RuntimeError:
                            self.release_current_email(uid)
                else:
                    self.log(f"No new emails - {self.current_time} {self.current_date}", tag="no_new_emails", write=False)
                    self.check_labels(["Need_Print", "Errors"], self.imap)
                    self.pause_event.wait(timeout=max(1, int(config.get("INBOX_CYCLE_TIME", 30))))

                cycle_count += 1
                cycle_time = max(1, int(config.get("INBOX_CYCLE_TIME", 30)))
                reconnect_time = max(cycle_time, int(config.get("RECONNECT_TIME", 3600)))
                reconnect_cycles = max(1, ceil(reconnect_time / cycle_time))
                if cycle_count >= reconnect_cycles:
                    self.reconnect()
                    cycle_count = 0
            except self.IMAPUnavailable:
                self.reconnect()
            except Exception as e:
                self.log(f"An error occurred while searching the inbox: {e}\n{traceback.format_exc()}", tag="red", send_email=True)
                self.reconnect()

        try:
            self.disconnect(self.imap)
        except Exception:
            pass
        if self.logging_out:
            self.logging_out = False
            self.ui(self.start_button.config, state=tk.NORMAL)
            self.ui(self.testing_button.config, state=tk.NORMAL)
            self.ui(self.away_mode_button.config, state=tk.NORMAL)
            self.ui(self._refresh_ui_indicators)

    def process_email(self, mail):
        subject = ""
        imap = None
        try:
            imap = self.safe_imap(self.connect, log=False, primary=False, use_lock=False)
            msg = self.get_msg(mail, "INBOX", imap)
            if msg is None:
                self.log(f"Email not found (process_email): {mail}", tag="red", send_email=True)
                self.release_current_email(mail)
                return
            subject = msg.get("Subject", "")
            sender_email = email.utils.parseaddr(msg.get("From", ""))[1]
            if not sender_email.lower().endswith(str(self.EMAIL_ENDING).lower()):
                self.move_email(mail, "Not_Invoices", "INBOX", imap)
                self.release_current_email(mail)
                return

            has_attachment = any(
                part.get_content_disposition() == "attachment"
                and part.get_filename() is not None
                and part.get_filename().lower().endswith(".pdf")
                for part in msg.walk()
            )
            if not has_attachment:
                label = str(config.get("NO_PDF_LABEL", "Not_Invoices") or "Not_Invoices")
                self.log(f"No PDF attachment in '{subject}', moving to {label}.", tag="orange")
                self.move_email(mail, label, "INBOX", imap)
                self.release_current_email(mail)
                return
            self.handle_attachments(mail, imap, msg, subject, sender_email)
        except self.IMAPUnavailable as e:
            self.log(f"IMAP unavailable while processing '{subject or mail}': {e}", tag="red")
            self.release_current_email(mail)
        except Exception as e:
            self.log(f"An error occurred while processing an email: {e}\n{traceback.format_exc()}", tag="red", send_email=True)
            try:
                if imap:
                    self.move_email(mail, "Errors", "INBOX", imap)
            except Exception:
                pass
            self.release_current_email(mail)
        finally:
            if imap:
                try:
                    imap.logout()
                except Exception:
                    pass

    def release_current_email(self, mail):
        with self.current_emails_lock:
            self.current_emails.discard(str(mail))

    @staticmethod
    def sha256_bytes(data):
        return hashlib.sha256(data).hexdigest()

    @staticmethod
    def sha256_file(filepath):
        digest = hashlib.sha256()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def get_document_history(self, document_hash):
        try:
            with sqlite3.connect(self.ARCHIVE_DB) as db:
                row = db.execute(
                    "SELECT sha256, is_invoice, saved, printed, final_name, filepath, first_seen, last_seen "
                    "FROM document_history WHERE sha256 = ?", (document_hash,)
                ).fetchone()
            if not row:
                return None
            keys = ("sha256", "is_invoice", "saved", "printed", "final_name", "filepath", "first_seen", "last_seen")
            return dict(zip(keys, row))
        except sqlite3.Error as exc:
            self.log(f"Could not query document history: {exc}", tag="orange")
            return None

    def record_document_history(self, document_hash, is_invoice=True, saved=False, printed=False,
                                final_name=None, filepath=None):
        if not document_hash:
            return
        now = datetime.now().isoformat(timespec="seconds")
        try:
            with sqlite3.connect(self.ARCHIVE_DB) as db:
                existing = db.execute(
                    "SELECT first_seen, saved, printed, final_name, filepath FROM document_history WHERE sha256 = ?",
                    (document_hash,),
                ).fetchone()
                first_seen = existing[0] if existing else now
                old_saved = bool(existing[1]) if existing else False
                old_printed = bool(existing[2]) if existing else False
                old_name = existing[3] if existing else None
                old_path = existing[4] if existing else None
                db.execute(
                    "INSERT OR REPLACE INTO document_history "
                    "(sha256, is_invoice, saved, printed, final_name, filepath, first_seen, last_seen) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        document_hash, int(bool(is_invoice)), int(bool(saved or old_saved)),
                        int(bool(printed or old_printed)), final_name or old_name, filepath or old_path,
                        first_seen, now,
                    ),
                )
                db.commit()
        except sqlite3.Error as exc:
            self.log(f"Could not update document history: {exc}", tag="orange")

    def recover_or_skip_duplicate(self, document_hash, mail, subject, filename):
        history = self.get_document_history(document_hash)
        if not history:
            return False
        is_invoice = bool(history["is_invoice"])
        self.valid_invoice_flags[mail] = self.valid_invoice_flags.get(mail, False) or is_invoice
        filepath = history.get("filepath")
        if is_invoice and history.get("saved") and not history.get("printed") and filepath and os.path.exists(filepath):
            self.log(f"Recovered previously saved-but-unprinted invoice {os.path.basename(filepath)}.", tag="yellow")
            printed = self.print_invoice(filepath)
            if printed:
                self.record_document_history(document_hash, True, True, True, os.path.basename(filepath), filepath)
        else:
            self.log(f"Duplicate attachment skipped: {filename}", tag="yellow")
        self.gui_queue.put((0, "NEW", mail, subject, f"Duplicate - {filename}", filepath or ""))
        self.gui_queue.put((0, "STATUS", f"{mail}_Duplicate - {filename}", "duplicate", filepath or ""))
        return True

    def resolve_destination(self, desired_path, document_hash=None):
        desired_path = os.path.abspath(desired_path)
        if not os.path.exists(desired_path):
            return desired_path, False
        try:
            existing_hash = self.sha256_file(desired_path)
            if document_hash and existing_hash == document_hash:
                return desired_path, True
        except OSError:
            pass
        base, ext = os.path.splitext(desired_path)
        suffix = (document_hash or hashlib.sha256(str(time.time_ns()).encode()).hexdigest())[:8]
        conflict = f"{base}_CONFLICT_{suffix}{ext}"
        counter = 2
        while os.path.exists(conflict):
            conflict = f"{base}_CONFLICT_{suffix}_{counter}{ext}"
            counter += 1
        self.log(
            f"Document identity collision: {os.path.basename(desired_path)} already exists with different content; "
            f"saving as {os.path.basename(conflict)}.", tag="orange", send_email=True,
        )
        return conflict, False

    def update_vendor_affixes(self, vendor, prefix=None, postfix=None):
        """Persist prefix/postfix values learned while creating a template."""
        vendor = str(vendor or "").strip()
        if not vendor:
            return

        def update_rule(setting_key, value):
            value = str(value or "").strip()
            raw = str(config.get(setting_key, "") or "")
            rules = []
            found = False
            for item in [part.strip() for part in raw.split(",") if part.strip()]:
                if ":" not in item:
                    rules.append(item)
                    continue
                name, current = item.split(":", 1)
                if name.strip().casefold() == vendor.casefold():
                    found = True
                    if value:
                        rules.append(f"{vendor}:{value}")
                    # Blank text removes an existing rule for this vendor.
                else:
                    rules.append(item)
            if not found and value:
                rules.append(f"{vendor}:{value}")
            config[setting_key] = ", ".join(rules)

        if prefix is not None:
            update_rule("PREFIX_VENDORS", prefix)
        if postfix is not None:
            update_rule("POSTFIX_VENDORS", postfix)
        with open(self.CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
        try:
            self.rectangulator_handler.refresh_config()
        except Exception:
            pass

    def open_folder(self, folder):
        try:
            folder = str(folder or "").strip()
            if not folder:
                return
            os.makedirs(folder, exist_ok=True)
            os.startfile(folder)
        except Exception as exc:
            messagebox.showerror("Open Folder", str(exc))

    def safe_pdf_name(self, filename):  # Sanitize filename to prevent issues
        base = os.path.basename(filename or "invoice.pdf")
        base = re.sub(r"[^\w.\- ]", "_", base).strip()
        if not base.lower().endswith(".pdf"):
            base += ".pdf"
        return base or "invoice.pdf"

    def handle_attachments(self, mail, imap, msg, subject, sender_email):
        if msg is None:
            self.log(f"Email not found (handle_attachments): {mail}", tag="red", send_email=True)
            self.release_current_email(mail)
            return
        seen_attachment_names = set()
        pending = []  # filename, filepath, sha256, ocr_mode

        for part in msg.walk():
            raw_name = part.get_filename()
            if (not raw_name or raw_name in seen_attachment_names or
                    part.get_content_disposition() is None or not raw_name.lower().endswith(".pdf")):
                continue
            seen_attachment_names.add(raw_name)
            filename = self.safe_pdf_name(raw_name)
            if self.TESTING == "test":
                filename = f"Test_{filename}"
            filepath = os.path.join(self.INVOICE_FOLDER, filename)
            if os.path.commonpath([os.path.abspath(filepath), os.path.abspath(self.INVOICE_FOLDER)]) != os.path.abspath(self.INVOICE_FOLDER):
                raise ValueError("Refusing to write outside the invoice folder")

            attachment = part.get_payload(decode=True)
            if not attachment:
                self.log(f"Empty PDF attachment '{filename}'", tag="orange")
                continue
            document_hash = self.sha256_bytes(attachment)
            if self.recover_or_skip_duplicate(document_hash, mail, subject, filename):
                continue

            if os.path.exists(filepath):
                stem, ext = os.path.splitext(filename)
                filepath = os.path.join(self.INVOICE_FOLDER, f"{stem}_incoming_{document_hash[:8]}{ext}")
                filename = os.path.basename(filepath)
            filepath = re.sub(r'[\n\r\t]', ' ', filepath)
            with open(filepath, "wb") as file:
                file.write(attachment)

            has_text = self.rectangulator_handler.document_has_embedded_text(filepath)
            if not has_text:
                self.log(f"No embedded text found in {filename}; routing to human-confirmed OCR.", tag="yellow")
                pending.append((filename, filepath, document_hash, True))
                self.root.after(0, self.flash_taskbar)
                continue

            template_exists = self.rectangulator_handler.check_templates(filepath, self.TEMPLATE_FOLDER, self.root)
            if template_exists:
                if template_exists[0] == "SPLIT_PDF":
                    self.log(f"Split invoice detected for {filename} -- {self.current_date} {self.current_time}", tag="yellow")
                    with fitz.open(filepath) as doc:
                        for i in range(len(doc)):
                            new_filename = f"{os.path.splitext(filename)[0]}_part{i+1}.pdf"
                            new_filepath = os.path.join(self.INVOICE_FOLDER, new_filename)
                            new_doc = fitz.open()
                            new_doc.insert_pdf(doc, from_page=i, to_page=i)
                            new_doc.save(new_filepath)
                            new_doc.close()
                            part_hash = self.sha256_file(new_filepath)
                            pending.append((new_filename, new_filepath, part_hash, False))
                    # Do not mark the multi-page source complete yet. If the program
                    # crashes while processing its split pages, the original email
                    # should be eligible for recovery on the next run.
                    os.remove(filepath)
                    continue

                desired = template_exists[0]
                new_filepath, duplicate_on_disk = self.resolve_destination(desired, document_hash)
                if duplicate_on_disk:
                    os.remove(filepath)
                    self.valid_invoice_flags[mail] = True
                    self.record_document_history(document_hash, True, True, True, os.path.basename(new_filepath), new_filepath)
                    continue
                os.rename(filepath, new_filepath)
                new_name = os.path.basename(new_filepath)
                self.gui_queue.put((0, "NEW", mail, subject, new_name, new_filepath))
                self.gui_queue.put((0, "STATUS", f"{mail}_{new_name}", "saved", new_name, new_filepath))
                self.valid_invoice_flags[mail] = True
                self.record_document_history(document_hash, True, True, False, new_name, new_filepath)
                printed = self.print_invoice(new_filepath, f"{mail}_{new_name}")
                if printed:
                    self.record_document_history(document_hash, True, True, True, new_name, new_filepath)
            else:
                self.root.after(0, self.flash_taskbar)
                pending.append((filename, filepath, document_hash, False))

        with self.state_lock:
            self.remaining_pdfs[mail] = len(pending)
        mode = self.queue_testing_mode(subject, sender_email)
        for queued_name, queued_path, document_hash, ocr_mode in pending:
            self.add_to_queue(mail, subject, queued_name, queued_path, testing=mode,
                              document_hash=document_hash, ocr_mode=ocr_mode)

        if not pending:
            with self.state_lock:
                self.remaining_pdfs.pop(mail, None)
                valid = self.valid_invoice_flags.pop(mail, False)
            label = "Invoices" if valid else "Not_Invoices"
            self.move_email(mail, label, "INBOX", self.imap)
            self.release_current_email(mail)

    def queue_testing_mode(self, subject, sender_email):  # Which testing flag add_to_queue should get
        if subject == "Test":
            return "test"
        elif self.TESTING:
            return True
        elif sender_email == config["SCANNER_EMAIL"]:
            return "scanner"
        return False

    def add_to_queue(self, mail, subject, filename, filepath, testing=False, document_hash=None, ocr_mode=False):
        try:
            self.gui_queue.put((1, "NEW", mail, subject, filename, filepath))
            self.gui_queue.put((2, "RECTANGULATE", mail, filename, filepath, testing, document_hash, bool(ocr_mode)))
        except Exception as e:
            self.log(f"An error occurred while processing the queue: {e}", tag="red", send_email=True)

    def process_queue(self):
        try:
            if not self.gui_queue.empty() and not self.gui_busy:
                task = self.gui_queue.get()[1:]
                task_type = task[0]
                if task_type == "NEW":
                    mail, subject, filename, filepath = task[1:5]
                    iid = f"{mail}_{filename}"
                    if iid not in self.inbox.get_children():
                        self.inbox.insert(
                            "", "end", iid=iid,
                            values=(subject, self.current_date, filename, "No", "No", "", filepath),
                            tags=("pending",),
                        )
                        self._update_queue_count()
                elif task_type == "STATUS":
                    _id, status = task[1:3]
                    if _id not in self.inbox.get_children():
                        return
                    item = self.inbox.item(_id)
                    values = list(item["values"])
                    if status == "saved":
                        filename, filepath = task[3:5]
                        values[2], values[6] = filename, filepath
                        if filename != "Not Invoice":
                            values[3] = "Yes"
                        self.inbox.item(_id, tags=("finished",))
                    elif status == "printed":
                        values[4] = "Yes"
                    elif status == "duplicate":
                        filepath = task[3] if len(task) > 3 else values[6]
                        values[3] = "Duplicate"
                        values[6] = filepath
                        self.inbox.item(_id, tags=("finished",))
                    elif status.startswith("Error"):
                        values[5] = status
                        self.inbox.item(_id, tags=("error",))
                    self.inbox.item(_id, values=values)
                elif task_type == "REMOVE":
                    _id = task[1]
                    if _id in self.inbox.get_children():
                        self.inbox.delete(_id)
                        self._update_queue_count()
                elif task_type == "RECTANGULATE":
                    self.gui_busy = True
                    mail, filename, filepath, testing, document_hash, ocr_mode = task[1:7]
                    self.root.after(
                        0, lambda m=mail, n=filename, p=filepath, t=testing, h=document_hash, o=ocr_mode:
                        self.handle_rectangulator(m, n, p, t, h, o)
                    )
        finally:
            self.root.after(100, self.process_queue)

    def handle_rectangulator(self, mail, filename, filepath, testing, document_hash=None, ocr_mode=False):
        counted = False
        inbox_item_id = f"{mail}_{filename}"
        try:
            # Native text templates may have been created while this invoice waited
            # in the UI queue. OCR templates never bypass human review.
            if not ocr_mode and testing is False:
                template_exists = self.rectangulator_handler.check_templates(filepath, self.TEMPLATE_FOLDER, self.root)
                if template_exists and template_exists[0] != "SPLIT_PDF":
                    desired = template_exists[0]
                    new_filepath, duplicate_on_disk = self.resolve_destination(desired, document_hash)
                    if duplicate_on_disk:
                        if os.path.exists(filepath):
                            os.remove(filepath)
                    else:
                        os.rename(filepath, new_filepath)
                    new_name = os.path.basename(new_filepath)
                    self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", new_name, new_filepath))
                    self.valid_invoice_flags[mail] = True
                    self.record_document_history(document_hash, True, True, False, new_name, new_filepath)
                    printed = True if duplicate_on_disk else self.print_invoice(new_filepath, inbox_item_id)
                    if printed:
                        self.record_document_history(document_hash, True, True, True, new_name, new_filepath)
                    self.finish_pdf(mail)
                    counted = True
                    return

            return_list = self.rectangulator_handler.rectangulate(
                filename, filepath, self, self.TEMPLATE_FOLDER, testing,
                ocr_mode=ocr_mode, ocr_template_folder=self.OCR_TEMPLATE_FOLDER,
                statement_folder=self.STATEMENT_FOLDER,
                statement_template_folder=self.STATEMENT_TEMPLATE_FOLDER,
            )
            print(f"-rectangulator returned: {return_list}")

            if self.AWAY_MODE:
                self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", f"Away Mode - {filename}", "None"))
                with self.state_lock:
                    self.remaining_pdfs.pop(mail, None)
                    self.valid_invoice_flags.pop(mail, None)
                counted = True
                return

            if return_list == [] or not return_list or return_list[0] is None:
                if not str(mail).startswith("REPROCESS_"):
                    self.move_email(mail, "Errors", "INBOX", self.imap)
                if os.path.exists(filepath):
                    os.remove(filepath)
                self.log(f"Failed to process '{filename}', moved to Error label, not printed", tag="red", send_email=True)
                self.gui_queue.put((0, "STATUS", inbox_item_id, "Error Rectangulating"))
                with self.state_lock:
                    self.remaining_pdfs.pop(mail, None)
                    self.valid_invoice_flags.pop(mail, None)
                counted = True
                self.release_current_email(mail)
                return

            if return_list[0] == "test_email":
                self.log("Test complete", tag="purple")
                self.move_email(mail, "Test_Email", "INBOX", self.imap)
                if os.path.exists(filepath):
                    os.remove(filepath)
                self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", "Test Email", "None"))
                with self.state_lock:
                    self.remaining_pdfs.pop(mail, None)
                    self.valid_invoice_flags.pop(mail, None)
                counted = True
                self.release_current_email(mail)
                return

            if return_list[0] == "statement":
                _, new_filepath, should_print = return_list
                # Scanner/MFP submissions are saved but intentionally not printed.
                if testing is True or testing == "scanner":
                    should_print = False
                new_filepath, duplicate_on_disk = self.resolve_destination(new_filepath, document_hash)
                if duplicate_on_disk:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                else:
                    os.makedirs(os.path.dirname(new_filepath), exist_ok=True)
                    os.rename(filepath, new_filepath)
                new_name = os.path.basename(new_filepath)
                self.valid_invoice_flags[mail] = True
                self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", new_name, new_filepath))
                self.log(f"Saved statement {new_name} -- {self.current_date} {self.current_time}", tag="lgreen")
                # A scanner/MFP statement intentionally has no print job. Mark
                # that print step complete in duplicate history so a restart or
                # duplicate email cannot unexpectedly print it later.
                print_complete = bool(duplicate_on_disk or not should_print)
                self.record_document_history(document_hash, True, True, print_complete, new_name, new_filepath)
                printed = duplicate_on_disk
                if should_print and not duplicate_on_disk:
                    printed = self.print_invoice(new_filepath, inbox_item_id)
                if printed:
                    self.record_document_history(document_hash, True, True, True, new_name, new_filepath)
                self.finish_pdf(mail)
                counted = True
                return

            new_filepath, should_print = return_list
            if new_filepath == "not_invoice":
                target_path, should_print, should_save = should_print
                if testing is True:
                    should_print = False
                if should_print:
                    self.print_invoice(filepath, inbox_item_id)
                if should_save:
                    target_path, duplicate_on_disk = self.resolve_destination(target_path, document_hash)
                    if duplicate_on_disk:
                        if os.path.exists(filepath):
                            os.remove(filepath)
                    else:
                        os.rename(filepath, target_path)
                    self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", "Not Invoice", target_path))
                    self.record_document_history(document_hash, False, True, bool(should_print), os.path.basename(target_path), target_path)
                else:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                    self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", "Not Invoice", "None"))
                    self.record_document_history(document_hash, False, False, bool(should_print), "Not Invoice", None)
                self.finish_pdf(mail)
                counted = True
                return

            self.valid_invoice_flags[mail] = True
            if testing is True:
                should_print = False
            new_filepath, duplicate_on_disk = self.resolve_destination(new_filepath, document_hash)
            if duplicate_on_disk:
                if os.path.exists(filepath):
                    os.remove(filepath)
            else:
                os.rename(filepath, new_filepath)
            new_name = os.path.basename(new_filepath)
            self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", new_name, new_filepath))
            self.log(f"Created new invoice file {new_name} -- {self.current_date} {self.current_time}", tag="lgreen")
            self.record_document_history(document_hash, True, True, False, new_name, new_filepath)
            printed = duplicate_on_disk
            if should_print and not duplicate_on_disk:
                printed = self.print_invoice(new_filepath, inbox_item_id)
            if printed:
                self.record_document_history(document_hash, True, True, True, new_name, new_filepath)
            self.finish_pdf(mail)
            counted = True
        except Exception as e:
            self.log(f"Rectangulator failed for '{filename}': {e}\n{traceback.format_exc()}", tag="red", send_email=True)
        finally:
            if not counted:
                self.finish_pdf(mail)
            self.gui_busy = False
            try:
                self.root.after(25, self._show_preview_idle)
            except tk.TclError:
                pass

    def finish_pdf(self, mail):
        with self.state_lock:
            if mail not in self.remaining_pdfs:
                return
            self.remaining_pdfs[mail] -= 1
            if self.remaining_pdfs[mail] > 0:
                return
            del self.remaining_pdfs[mail]
            valid = self.valid_invoice_flags.pop(mail, False)

        if str(mail).startswith("REPROCESS_"):
            self.release_current_email(mail)
            return
        label = "Invoices" if valid else "Not_Invoices"
        self.move_email(mail, label, "INBOX", self.imap)
        self.release_current_email(mail)

    def remove_inbox_item(self, log=True):  # Removes inbox item on double click and adds to archive
        # Only if item has finished tag
        item = self.inbox.selection()  # Get selected item
        if not item or "finished" not in self.inbox.item(item[0], "tags"):
            return
        
        _id = item[0]
        values = self.inbox.item(_id, "values")  # Get values of selected item
        new_id = f"{_id}_{int(time.time())}" # create new unique id for archive item

        if values[2] == "Test Email":  # If item is test email, don't archive
            self.inbox.delete(_id)
            self._update_queue_count()
            return
        
        self.archive.insert(
            "", "end",
            iid=new_id,
            values=(values[0], values[1], values[2], values[3], values[4], values[5], values[6]),
            tags=("default",))
        self.save_archive((new_id, values[0], values[1], values[2], values[3], values[4], values[5], values[6]))  # Save to archive
        self.inbox.delete(_id)  # Remove item from inbox
        self._update_queue_count()
        if log:
            self.log(f"Archived inbox item {_id} as {new_id}.", tag="blue")

    def remove_archive_item(self, event):  # Removes archive item on double right click
        item = self.archive.selection()  # Get selected item
        if not item:
            return
        
        _id = item[0]
        self.archive.delete(_id)  # Remove item from archive tree
        self.db.execute("DELETE FROM archive WHERE id = ?", (_id,))  # Remove from database
        self.db.commit() 
        self.log(f"Removed archive item {_id} from archive.", tag="blue")  # Log removal

    def reprocess_archive_item(self):
        item = self.archive.selection()
        if not item:
            messagebox.showinfo("Reprocess", "Select an archived item first.")
            return
        source = self.archive.item(item[0], "values")[-1]
        if not source or not os.path.exists(source):
            messagebox.showerror("Reprocess", f"File not found: {source}")
            return
        try:
            os.makedirs(self.INVOICE_FOLDER, exist_ok=True)
            token = str(time.time_ns())
            temp_name = f"Reprocess_{token}_{os.path.basename(source)}"
            temp_path = os.path.join(self.INVOICE_FOLDER, temp_name)
            shutil.copy2(source, temp_path)
            mail = f"REPROCESS_{token}"
            document_hash = None  # explicit reprocesses intentionally bypass duplicate suppression
            ocr_mode = not self.rectangulator_handler.document_has_embedded_text(temp_path)
            with self.state_lock:
                self.remaining_pdfs[mail] = 1
                self.valid_invoice_flags[mail] = False
            self.add_to_queue(mail, "Archive Reprocess", temp_name, temp_path,
                              testing="reprocess", document_hash=document_hash, ocr_mode=ocr_mode)
            self.log(f"Queued archive file for reprocessing: {os.path.basename(source)}", tag="blue")
        except Exception as exc:
            messagebox.showerror("Reprocess", str(exc))

    def print_archive_item(self, event): # Prints archive item on middle click
        item = self.archive.selection()
        if not item:
            return
        
        _id = item[0]
        filepath = self.archive.item(_id, "values")[-1]
        if not filepath or not os.path.exists(filepath):
            self.log(f"File not found for printing archive item {_id}, {filepath}", tag="red")
            return
        try:
            self.print_invoice(filepath)
        except Exception as e:
            self.log(f"Error printing archive item {_id}: {str(e)}", tag="red")

    def archive_all(self):  # Archives all inbox items
        self.log(f"Archiving {len(self.inbox.get_children())} items.", tag="blue")
        for item in self.inbox.get_children(): # select item and call remove
            self.inbox.selection_set(item)  # Select item
            self.remove_inbox_item(log=False)  # Call remove inbox item method

    def save_archive(self, record):
        self.db.execute("INSERT OR REPLACE INTO archive VALUES (?, ?, ?, ?, ?, ?, ?, ?)", record)
        self.db.commit()

    def load_archive(self):  # Loads archive from database
        values = []
        for row in self.db.execute("SELECT * FROM archive"):
            values.append(row)
        values.sort(key=lambda x: datetime.strptime(x[2], "%m-%d-%Y")  if x[2] else datetime(2000, 1, 1), reverse=True)
        for row in values:
            self.archive.insert("", "end", iid=row[0], values=row[1:])

    def open_archive_item(self, event):  # Opens archive item on double click
        item = self.archive.selection()
        if not item:
            return
        _id = item[0]
        filepath = self.archive.item(_id, "values")[-1]  # Get filepath from selected item
        if not filepath or not os.path.exists(filepath):
            self.log(f"File not found for archive item {_id}: {filepath}", tag="red", send_email=True)
            return
        try:
            os.startfile(filepath)  # Open the file
        except Exception as e:
            self.log(f"Error opening file {filepath}: {str(e)}", tag="red")

    def move_email(self, mail, label, og_label, imap):
        subject = "Unknown"
        try:
            guard = self.imap_lock if imap is getattr(self, "imap", None) else nullcontext()
            with guard:
                msg = self.get_msg(mail, og_label, imap)
                if msg:
                    subject = msg.get("Subject", "Unknown")
                imap.select(og_label)
                success = imap.uid("COPY", mail, label)
                if success[0] != "OK":
                    self.log(f"Error copying email '{subject}': {success[1]}", tag="red", send_email=True)
                    return False
                imap.uid("STORE", mail, "+FLAGS", "(\\Deleted)")
                imap.expunge()
            self.log(f"Moved email '{subject}' from {og_label} to {label}.", tag="blue")
            return True
        except Exception as e:
            self.log(f"Transfer failed for '{subject}': {e}\n{traceback.format_exc()}", tag="red", send_email=True)
            return False

    def send_email(self, body):  # Sends email to me
        sender_email = f"{self.username}.sndex@gmail.com"
        try:
            if self.TESTING:
                return

            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["Subject"] = "Alert"
            message["From"] = sender_email
            message["To"] = config["RECEIVER_EMAIL"]
            message.attach(MIMEText(body, "plain"))

            # Send the email using SMTP
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, self.password)
                server.sendmail(sender_email, config["RECEIVER_EMAIL"], message.as_string())
        except Exception as e:
            self.log(f"Error sending email - {str(e)}", tag="red")

    def get_subject(self, mail, label, imap):  # Get the message from the specified email
        try:
            with self.imap_lock:
                imap.select(label)
                result, data = imap.uid("FETCH", mail, "(BODY.PEEK[])")
                if result != "OK" or not data or not data[0]:
                    self.log(f"Error fetching email: {mail}", tag="red", send_email=True)
                    return None
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                return msg["Subject"]
        except Exception as e:
            self.log(f"Error getting subject: {str(e)}", tag="red", send_email=True)
            return "Unknown"

    def get_email(self, label, imap):  # Gets most recent email uid in label
        try:
            with self.imap_lock:
                imap.select(label)
                _, data = imap.uid("search", None, "ALL")
                uid = data[0].split()[-1]
                return uid
        except Exception as e:
            return None

    def print_invoice(self, filepath, inbox_item_id=None):
        try:
            printer_name = str(config.get("PRINTER_NAME", "") or "").strip()
            if printer_name:
                printers = [p[2] for p in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
                )]
                match = next((p for p in printers if p.casefold() == printer_name.casefold()), None)
                if not match:
                    raise RuntimeError(f"Configured printer '{printer_name}' is not installed/available")
                # printto asks the registered PDF application to target this printer.
                result = win32api.ShellExecute(0, "printto", filepath, f'"{match}"', ".", 0)
                destination = match
            else:
                result = win32api.ShellExecute(0, "print", filepath, None, ".", 0)
                try:
                    destination = win32print.GetDefaultPrinter()
                except Exception:
                    destination = "Windows default printer"
            if int(result) <= 32:
                raise OSError(f"ShellExecute returned {result}")
            self.log(f"Sent {os.path.basename(filepath)} to {destination}.", tag="lgreen")
            if inbox_item_id:
                self.gui_queue.put((0, "STATUS", inbox_item_id, "printed"))
            return True
        except Exception as e:
            self.log(f"Printing failed for {filepath}: {e}", tag="red", send_email=True)
            if inbox_item_id:
                self.gui_queue.put((0, "STATUS", inbox_item_id, "Error Printing"))
            return False

    def log(self, *args, tag=None, send_email=False, write=True, console=True):
        message = " ".join(str(a) for a in args)

        # Disk + network are thread-safe; do them here.
        if write:
            try:
                with open(config["LOG_FILE"], "a", encoding="utf-8") as f:
                    f.write(message + "\n")
            except OSError as e:
                print(f"-log file error: {e}")
        if send_email and tag == "red":
            threading.Thread(target=self.send_email, args=(message,), daemon=True).start()

        # Widget access must run on the main thread.
        if not console:
            return
        try:
            self.root.after(0, self._log_to_widget, message, tag)
        except (RuntimeError, tk.TclError):
            pass

    def ui(self, fn, *args, **kwargs):
        try:
            self.root.after(0, lambda: fn(*args, **kwargs))
        except (RuntimeError, tk.TclError):
            pass

    def _log_to_widget(self, message, tag):
        if self.window_closed:
            return
        try:
            if tag in ("no_new_emails", "label_error"):
                self.remove_messages(message)
            else:
                print(f"-{message}")
            self.log_text_widget.insert(tk.END, message + "\n", (tag, "default"))
            if self.log_text_widget.yview()[1] > 0.75:
                self.log_text_widget.yview_moveto(1)
        except tk.TclError:
            pass

    def check_labels(self, labels, imap):  # Checks for emails that need to be looked at in labels
        # If passed one label, returns email uids, otherwise just logs the number of emails in each label
        for label in labels:
            try:
                with self.imap_lock:
                    # Check if any emails in specified label
                    imap.select(label)
                    _, data = imap.uid("search", None, "ALL")
                    email_ids = data[0].split()

                    # Alert user if there are emails
                    if len(labels) == 1: # When resolve button pressed
                        return email_ids
                    elif len(email_ids) > 0:
                        self.log(f"{len(email_ids)} emails in {label} - {self.current_time} {self.current_date}", tag="label_error", write=False)
                        if label == "Errors":
                            self.ui(self._set_button_attention, self.errors_button, True)
                        elif label == "Need_Print":
                            self.ui(self._set_button_attention, self.print_errors_button, True)
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, TimeoutError, socket.gaierror, ssl.SSLError, OSError):
                raise
            except Exception as e:
                # A GUI/configuration problem should be logged, but it should not be
                # disguised as an IMAP disconnect and force a Gmail reconnect.
                self.log(f"An error occurred while checking the label: {str(e)}", tag="red", send_email=True)

    def get_msg(self, mail, label, imap):
        try:
            guard = self.imap_lock if imap is getattr(self, "imap", None) else nullcontext()
            with guard:
                imap.select(label)
                result, data = imap.uid("FETCH", mail, "(BODY.PEEK[])")
                if result != "OK" or not data or not data[0]:
                    return None
                return email.message_from_bytes(data[0][1])
        except Exception as e:
            self.log(f"Error getting email message {mail}: {e}", tag="red")
            return None

    def remove_messages(self, message):  # Removes no_new_emails messages
        message = message[:-22]  # cuts out the date-time

        # Searches for every no_new_emails message then deletes it
        index = self.log_text_widget.search(message, "1.0", tk.END)
        while index:
            self.log_text_widget.delete(index, f"{index}+{len(message) + 23}c")  # +1 for new line, +22 for date-time
            index = self.log_text_widget.search(message, "1.0", tk.END)
            self.root.update()

    def resolve_errors(self):  # Moves error emails back to inbox
        try:
            self.log(f"Attempting to resolve errors.", tag="yellow")
            # Get emails in error label
            email_ids = self.check_labels(["Errors"], self.imap)

            if len(email_ids) == 0:
                self.log(f"No errors to resolve.", tag="yellow")
                return

            # Move emails back to inbox
            for email_id in email_ids:
                self.move_email(email_id, "INBOX", "Errors", self.imap)
            self._set_button_attention(self.errors_button, False)
        except Exception as e:
            self.log(f"Error resolving errors: {str(e)}", tag="red", send_email=True)

    def resolve_prints(self):  # Moves unprinted invoices back to inbox
        try:
            self.log(f"Attempting to resolve unprinted invoices.", tag="yellow")
            # Get emails in Need_Print label
            email_ids = self.check_labels(["Need_Print"], self.imap)

            if len(email_ids) == 0:
                self.log(f"No unprinted invoices to resolve.", tag="yellow")
                return

            # Move emails back to inbox
            for email_id in email_ids:
                self.move_email(email_id, "INBOX", "Need_Print", self.imap)
            self._set_button_attention(self.print_errors_button, False)
        except Exception as e:
            self.log(f"Error resolving unprinted invoices: {str(e)}", tag="red", send_email=True)

    def pause_processing(self):  # Pauses processing
        self.log("Processing paused.", tag="yellow")
        self.pause_button.config(text="▶  Resume", command=self.resume_processing)
        self.errors_button.config(state=tk.DISABLED)
        self.print_errors_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)
        self.pause_event.set()
        self._refresh_ui_indicators()

    def resume_processing(self):  # Resumes processing
        self.log("Processing resumed.", tag="yellow")
        self.pause_button.config(text="Ⅱ  Pause", command=self.pause_processing)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.test_inbox_button.config(state=tk.NORMAL)
        self.pause_event.clear()
        self._refresh_ui_indicators()

    def restart_processing(self):  # Restarts processing
        self.log("Restarting...", tag="yellow")
        self.processor_running = False
        self.pause_event.clear()
        try:
            self.disconnect(getattr(self, "imap", None), log=False)
        except Exception:
            pass
        if self.processor_thread and self.processor_thread.is_alive():
            self.processor_thread.join(timeout=2)
        if self.email_executor is not None:
            self.email_executor.shutdown(wait=False, cancel_futures=False)
            self.email_executor = None
        self.main()

    def logout(self, reconnect=False):  # Logs out
        self.log("Logging out...", tag="yellow")
        self.ui(self.button_logout)
        self.pause_event.set()
        self.processor_running = False
        self.logging_out = True
        self.current_emails.clear()  # clear current emails
        self._refresh_ui_indicators()
        if self.email_executor is not None:
            self.email_executor.shutdown(wait=False, cancel_futures=False)
            self.email_executor = None

        if reconnect:
            # wait a few seconds then reconnect
            time.sleep(5)
            self.main()  # restart main loop

    def toggle_testing(self):  # Toggles testing mode
        self.TESTING = not self.TESTING
        self._refresh_ui_indicators()

    def toggle_away_mode(self):  # Toggles away mode, always prints an saves invoice
        self.AWAY_MODE = not self.AWAY_MODE
        self._refresh_ui_indicators()

    def test_rectangulator(self):
        self.log("Testing Rectangulator...", tag="yellow")
        test_path = config["TEST_INVOICE"]
        ocr_mode = not self.rectangulator_handler.document_has_embedded_text(test_path)
        template_folder = config.get("TEST_OCR_TEMPLATE_FOLDER") if ocr_mode else config["TEST_TEMPLATE_FOLDER"]
        return_list = self.rectangulator_handler.rectangulate(
            "Testing Rectangulator", test_path, self,
            config["TEST_TEMPLATE_FOLDER"], True,
            ocr_mode=ocr_mode,
            ocr_template_folder=config.get("TEST_OCR_TEMPLATE_FOLDER", template_folder),
        )
        if return_list:
            self.log(f"Test complete - result: {return_list}", tag="purple")
        self.log("Testing complete.", tag="yellow")

    def test_inbox(self):  # Sends test email to inbox, won't be printed or downloaded
        self.log("Sending test email to inbox", tag="yellow")
        mail = self.get_email("Test_Email", self.imap)
        if mail is None:
            self.log("Test email missing", tag="orange")
            return
        self.move_email(mail, "INBOX", "Test_Email", self.imap)

    def reconnect(self):
        try:
            self.disconnect(getattr(self, "imap", None), log=False)
        except Exception:
            pass
        self.imap = None
        self.connected = False
        while self.processor_running and self.imap is None:
            try:
                self.imap = self.safe_imap(self.connect, primary=True, retries=2, log_errors=False)
            except self.IMAPUnavailable:
                self.imap = None
            if self.imap:
                self.log(f"Reconnected to {self.username} -- {self.current_time} {self.current_date}", tag="green")
                self.update_crash_counter_label()
                return
            self.log("Reconnect failed, trying again in 30 seconds...", tag="orange")
            time.sleep(30)

    def on_program_exit(self):  # Runs when program is closed, disconnects and closes window
        self.log("Disconnecting...", tag="red")
        self.window_closed = True
        self.save_crash_counter()
        self.archive_all()  # Archive all inbox items

        # Disconnect imaps if running
        if self.processor_thread:
            self.processor_running = False
            self.pause_event.set()
            self.processor_thread.join(timeout=1)  # Wait for thread to finish

        if self.email_executor is not None:
            self.email_executor.shutdown(wait=False, cancel_futures=False)
            self.email_executor = None
        # Destroys tkinter window
        self.root.destroy()

    def flash_taskbar(self):  # Flash icon in taskbar
        # Code from stack overflow
        hwnd_int = int(self.root.frame(), base=16)
        win32gui.FlashWindow(hwnd_int, 0)

    def load_crash_counter(self): # Sets date variable from config
        try:
            last_crash_date = config.get("LAST_CRASH_DATE")
            self.last_crash_date = last_crash_date if last_crash_date else str(datetime.now().strftime("%Y-%m-%d"))
        except Exception as e:
            print(f"-No crash date found {str(e)}")
            self.last_crash_date = str(datetime.now().strftime("%Y-%m-%d"))
            self.save_crash_counter()

    def save_crash_counter(self): # Overwrites date in config
        try:
            config["LAST_CRASH_DATE"] = self.last_crash_date
            with open(os.path.join(BASE_DIR, "config.json"), "w") as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"-what happened {str(e)}")

    def reset_crash_counter(self): # Reset date variable and updates label
        self.last_crash_date = str(datetime.now().strftime("%Y-%m-%d"))
        self.save_crash_counter()
        self.update_crash_counter_label()

    def update_crash_counter_label(self):
        self.ui(self.days_without_crashing.set, f"Days without crashing: {(datetime.today() - datetime.strptime(self.last_crash_date, '%Y-%m-%d')).days}")

    @property
    def current_time(self):
        return time.strftime("%H:%M:%S", time.localtime())

    @property
    def current_date(self):
        return time.strftime("%m-%d-%Y", time.localtime())