import Rectangulator
import tkinter as tk
from tkinter import ttk, messagebox
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
_CONFIG_DEFAULTS = {
    "OCR_TEMPLATE_FOLDER": os.path.join(_template_parent, "OCR_Templates"),
    "TEST_OCR_TEMPLATE_FOLDER": os.path.join(_test_template_parent, "OCR_Templates"),
    "OCR_FUZZY_THRESHOLD": 0.72,
    "OCR_DPI": 250,
    "OCR_LANGUAGE": "eng",
    "TESSDATA_PREFIX": "",
    "PRINTER_NAME": "",                 # blank = Windows default printer
    "EMAIL_WORKERS": 3,
    "NO_PDF_LABEL": "Not_Invoices",
    "MIN_EMBEDDED_TEXT_CHARS": 40,
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


class EmailProcessor:

    # CONSTANTS
    CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
    ICON_PATH = os.path.join(BASE_DIR, "Hotpot.ico")
    TEMPLATE_FOLDER = config["TEMPLATE_FOLDER"]
    OCR_TEMPLATE_FOLDER = config["OCR_TEMPLATE_FOLDER"]
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

            # GUI
            self.root = tk.Tk()
            self.root.iconbitmap(self.ICON_PATH)
            # Size relative to the actual screen, then start maximized so the
            # rectangulator pane gets as much vertical room as possible.
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            self.root.geometry(f"{int(screen_w * 0.95)}x{int(screen_h * 0.90)}+0+0")
            self.root.minsize(1100, 650)
            try:
                self.root.state("zoomed")  # Windows / most Linux WMs
            except tk.TclError:
                pass
            self.root.title(f"{username.upper()} Pewter")
            style = ttk.Style(self.root)
            style.theme_use("clam")

            # Notebook and tabs
            notebook = ttk.Notebook(self.root)
            notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            # Program tab
            program_tab = tk.Frame(notebook)
            notebook.add(program_tab, text="Pewter") 

            # Layout frames (left for console, right for rectangulator)
            self.alert_container = tk.Frame(program_tab, relief="raised") # container for alert popup
            self.alert_container.place(relx=0.5, rely=0.5, anchor=tk.CENTER)  
            self.alert_container.lower()  # hide initially
            # Draggable sash: grab the divider to trade console width for PDF width.
            # stretch="never" on the left pane means every pixel gained by
            # maximizing the window goes to the rectangulator canvas.
            self.main_paned = tk.PanedWindow(program_tab,
                                             orient=tk.HORIZONTAL,
                                             sashwidth=6,
                                             sashrelief="raised",
                                             opaqueresize=False)
            self.main_paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            left_frame = tk.Frame(self.main_paned)
            right_frame = tk.Frame(self.main_paned)
            self.main_paned.add(left_frame, minsize=420, stretch="never")
            self.main_paned.add(right_frame, minsize=520, stretch="always")

            # Button Frame
            button_frame = tk.Frame(left_frame)  # frame for buttons
            button_frame.pack(side=tk.TOP, fill=tk.X, pady=2)
            # GUI Buttons
            self.start_button = tk.Button( # start process button
                button_frame, text="Start",
                command=self.main)  
            self.start_button.pack(side=tk.LEFT, padx=1)

            self.pause_button = tk.Button(# pause button
                button_frame,
                text="Pause",
                command=self.pause_processing,
                state=tk.DISABLED,
            )  
            self.pause_button.pack(side=tk.LEFT, padx=1)

            self.logout_button = tk.Button(
                button_frame,
                text="Logout",
                command=self.logout,
                state=tk.DISABLED)  # logout button
            self.logout_button.pack(side=tk.LEFT, padx=1)

            self.errors_button = tk.Button( # resolve errors button
                button_frame,
                text="Resolve Errors",
                command=self.resolve_errors,
                state=tk.DISABLED,
            )  #
            self.errors_button.pack(side=tk.LEFT, padx=1)

            self.print_errors_button = tk.Button( # resolve unprinted invoices button
                button_frame,
                text="Resolve Prints",
                command=self.resolve_prints,
                state=tk.DISABLED,
            )  
            self.print_errors_button.pack(side=tk.LEFT, padx=1)

            self.clear_button = tk.Button( # clear button
                button_frame,
                text="Clear",
                command=lambda: self.log_text_widget.delete("1.0", tk.END),
                state=tk.NORMAL,
            )  
            self.clear_button.pack(side=tk.LEFT, padx=1)

            self.testing_button = tk.Button( # testing button
                button_frame,
                text="Testing",
                command=self.toggle_testing,
                state=tk.NORMAL,
                bg="#FFCCCC",
                fg="black",
            )  
            self.testing_button.pack(side=tk.LEFT, padx=1)

            self.away_mode_button = tk.Button( # away mode button
                button_frame,
                text="Away Mode",
                command=self.toggle_away_mode,
                state=tk.NORMAL,
                bg="#FFCCCC",
                fg="black",
            )  
            self.away_mode_button.pack(side=tk.LEFT, padx=1)

            self.test_rectangulator_button = tk.Button( # test rectangulator button
                button_frame,
                text="Test Rectangulator",
                command=self.test_rectangulator,
                state=tk.NORMAL,
            )  
            self.test_rectangulator_button.pack(side=tk.LEFT, padx=1)

            self.test_inbox_button = tk.Button( # test inbox button
                button_frame,
                text="Test Inbox",
                command=self.test_inbox,
                state=tk.DISABLED,
            )  
            self.test_inbox_button.pack(side=tk.LEFT, padx=1)

            self.archive_all_button = tk.Button( # archive all button
                button_frame,
                text="Archive All",
                command=self.archive_all,
                state=tk.NORMAL,
            )
            self.archive_all_button.pack(side=tk.LEFT, padx=1)


            # Inbox / Log Frames (Left Frame)
            # Inbox Frame
            inbox_frame = tk.Frame(left_frame)  # frame for inbox treeview
            inbox_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            self.inbox = ttk.Treeview(inbox_frame,
                                      columns=("Subject", "Date", "Invoice", "Saved", "Printed", "Errors", "Filepath"),
                                      show="headings",
                                      height=15)
            self.inbox.column("Subject", width=150, anchor="center")
            self.inbox.column("Date", width=50, anchor="center")
            self.inbox.column("Invoice", width=100, anchor="center")
            self.inbox.column("Saved", width=30, anchor="center")
            self.inbox.column("Printed", width=30, anchor="center")
            self.inbox.column("Errors", width=60, anchor="center")
            self.inbox.column("Filepath", width=0, stretch=False)
            self.inbox.heading("Subject", text="Subject")
            self.inbox.heading("Date", text="Date")
            self.inbox.heading("Invoice", text="Invoice #")
            self.inbox.heading("Saved", text="Saved")
            self.inbox.heading("Printed", text="Printed")
            self.inbox.heading("Errors", text="Errors")
            self.inbox.heading("Filepath", text="")
            self.inbox.pack(fill=tk.BOTH, expand=True)
            self.inbox.bind("<Double-1>", self.remove_inbox_item)  # double click to remove inbox item
            self.inbox.tag_configure("pending", background="#FBFF1D")  # style for pending items
            self.inbox.tag_configure("finished", background="#68FF43")  # style for default items
            self.inbox.tag_configure("error", background="#FD4848")  # style for label errors
            
            # Log Frame
            log_frame = tk.Frame(left_frame)  # frame for log text widget
            log_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

            scrollbar = tk.Scrollbar(left_frame)  # scrollbar for log text widget
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.log_text_widget = tk.Text( # text box for logging
                left_frame,
                yscrollcommand=scrollbar.set,
                height=12,
                width=60,
                spacing1=4,
                padx=0,
                pady=0,
            )  
            self.log_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.configure(command=self.log_text_widget.yview)


            # Plot canvas for rectangulator
            self.figure = Figure(figsize=(9, 9.5), dpi=100)
            self.ax = self.figure.add_subplot(111)
            # Reclaim matplotlib's default margins; the axis is hidden anyway.
            # top=0.90 leaves room for the instruction labels drawn at 0.925-0.975,
            # bottom=0.11 for the text box / buttons drawn at 0.005-0.08.
            self.figure.subplots_adjust(left=0.05, right=0.95, top=0.90, bottom=0.11)
            self.ax.axis("off")
            self.canvas = FigureCanvasTkAgg(self.figure, master=right_frame)
            self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH,expand=True)

            self.gui_queue = queue.PriorityQueue()  # queue for all gui tasks
            self.gui_busy = False
            self.rectangulator_handler = Rectangulator.RectangulatorHandler(self, self.figure, self.ax)


            # Archive tab
            archive_tab = tk.Frame(notebook)
            notebook.add(archive_tab, text="Archive")
            archive_controls = tk.Frame(archive_tab)
            archive_controls.pack(side=tk.TOP, fill=tk.X, pady=2)
            tk.Button(archive_controls, text="Reprocess Selected", command=self.reprocess_archive_item).pack(side=tk.LEFT, padx=2)
            tk.Button(archive_controls, text="Open File", command=lambda: self.open_archive_item(None)).pack(side=tk.LEFT, padx=2)
            self.archive = ttk.Treeview(archive_tab,
                                        columns=("Subject", "Date", "Invoice", "Saved", "Printed", "Errors", "Filepath"),
                                        show="headings",
                                        height=15)
            self.archive.column("Subject", width=100, anchor="center")
            self.archive.column("Date", width=50, anchor="center")
            self.archive.column("Invoice", width=100, anchor="center")
            self.archive.column("Saved", width=50, anchor="center")
            self.archive.column("Printed", width=50, anchor="center")
            self.archive.column("Errors", width=100, anchor="center")
            self.archive.column("Filepath", width=300, anchor="center")
            self.archive.heading("Subject", text="Subject")
            self.archive.heading("Date", text="Date")
            self.archive.heading("Invoice", text="Invoice")
            self.archive.heading("Saved", text="Saved")
            self.archive.heading("Printed", text="Printed")
            self.archive.heading("Errors", text="Errors")
            self.archive.heading("Filepath", text="Filepath")
            self.archive.bind("<Double-1>", self.open_archive_item)  # double click to open archive item
            self.archive.bind("<Double-Button-3>", self.remove_archive_item)  # double right click to remove archive item
            self.archive.bind("<Button-2>", self.print_archive_item) # middle click to print archive item
            self.archive.pack(side="left", fill="both", expand=True)
            self.archive_scrollbar = ttk.Scrollbar(archive_tab,
                                                   orient="vertical",
                                                   command = self.archive.yview)
            self.archive_scrollbar.pack(side="right", fill="y")
            self.archive.configure(yscrollcommand=self.archive_scrollbar.set)

            # Load archive from database
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
            self.load_archive()  # Load archive from database
            

            # Console tab
            console_tab = tk.Frame(notebook)
            notebook.add(console_tab, text="Console")

            c_scrollbar = tk.Scrollbar(console_tab)  # scrollbar for console
            c_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            console_text = tk.Text(console_tab, yscrollcommand=c_scrollbar.set)
            console_text.pack(fill=tk.BOTH, expand=True)
            c_scrollbar.config(command=console_text.yview)


            # Settings tab
            settings_tab = tk.Frame(notebook)
            notebook.add(settings_tab, text="Settings")
            settings = {
                "APC_USER": tk.StringVar(value=config.get("APC_USER", "")),
                "LOG_FILE": tk.StringVar(value=config.get("LOG_FILE", "")),
                "INVOICE_FOLDER": tk.StringVar(value=config.get("INVOICE_FOLDER", "")),
                "TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEMPLATE_FOLDER", "")),
                "OCR_TEMPLATE_FOLDER": tk.StringVar(value=config.get("OCR_TEMPLATE_FOLDER", "")),
                "TEST_INVOICE_FOLDER": tk.StringVar(value=config.get("TEST_INVOICE_FOLDER", "")),
                "TEST_TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEST_TEMPLATE_FOLDER", "")),
                "TEST_OCR_TEMPLATE_FOLDER": tk.StringVar(value=config.get("TEST_OCR_TEMPLATE_FOLDER", "")),
                "TEST_INVOICE": tk.StringVar(value=config.get("TEST_INVOICE", "")),
                "INBOX_CYCLE_TIME": tk.IntVar(value=config.get("INBOX_CYCLE_TIME", 30)),
                "RECONNECT_TIME": tk.IntVar(value=config.get("RECONNECT_TIME", 3600)),
                "RECEIVER_EMAIL": tk.StringVar(value=config.get("RECEIVER_EMAIL", "")),
                "SCANNER_EMAIL": tk.StringVar(value=config.get("SCANNER_EMAIL", "")),
                "SPLIT_VENDORS": tk.StringVar(value=config.get("SPLIT_VENDORS", "")),
                "PREFIX_VENDORS": tk.StringVar(value=config.get("PREFIX_VENDORS", "")),
                "PRINTER_NAME": tk.StringVar(value=config.get("PRINTER_NAME", "")),
                "OCR_FUZZY_THRESHOLD": tk.DoubleVar(value=config.get("OCR_FUZZY_THRESHOLD", 0.72)),
                "OCR_DPI": tk.IntVar(value=config.get("OCR_DPI", 250)),
                "OCR_LANGUAGE": tk.StringVar(value=config.get("OCR_LANGUAGE", "eng")),
                "TESSDATA_PREFIX": tk.StringVar(value=config.get("TESSDATA_PREFIX", "")),
                "EMAIL_WORKERS": tk.IntVar(value=config.get("EMAIL_WORKERS", 3)),
                "NO_PDF_LABEL": tk.StringVar(value=config.get("NO_PDF_LABEL", "Not_Invoices")),
                "MIN_EMBEDDED_TEXT_CHARS": tk.IntVar(value=config.get("MIN_EMBEDDED_TEXT_CHARS", 40)),
            }
            for i, (key, var) in enumerate(settings.items()):
                display_key = "PRINTER_NAME (blank = Windows default)" if key == "PRINTER_NAME" else key
                label = tk.Label(settings_tab, text=display_key + ":")
                label.grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
                entry = tk.Entry(settings_tab, textvariable=var, width=100)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky=tk.W)

            def save_settings():
                # Saves settings to config.py
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
                    to_write = {k: v for k, v in config.items()}
                    with open(self.CONFIG_PATH, "w", encoding="utf-8") as f:
                        json.dump(to_write, f, indent=2)
                    self.TEMPLATE_FOLDER = config.get("TEMPLATE_FOLDER", self.TEMPLATE_FOLDER)
                    self.OCR_TEMPLATE_FOLDER = config.get("OCR_TEMPLATE_FOLDER", self.OCR_TEMPLATE_FOLDER)
                    self.INVOICE_FOLDER = config.get("INVOICE_FOLDER", self.INVOICE_FOLDER)
                    for folder_key in ("TEMPLATE_FOLDER", "OCR_TEMPLATE_FOLDER", "INVOICE_FOLDER"):
                        folder = str(config.get(folder_key, "") or "").strip()
                        if folder:
                            os.makedirs(folder, exist_ok=True)
                    self.rectangulator_handler.refresh_config()
                    self.refresh_template_manager()
                    messagebox.showinfo("Settings", "Settings saved successfully.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save settings: {e}")
            save_button = tk.Button(settings_tab, text="Save Settings", command=save_settings)
            save_button.grid(row=len(settings), column=0, columnspan=2, pady=10)
            tk.Button(settings_tab, text="Open Text Templates", command=lambda: self.open_folder(config.get("TEMPLATE_FOLDER", ""))).grid(row=len(settings)+1, column=0, pady=5)
            tk.Button(settings_tab, text="Open OCR Templates", command=lambda: self.open_folder(config.get("OCR_TEMPLATE_FOLDER", ""))).grid(row=len(settings)+1, column=1, pady=5, sticky=tk.W)

            # Template Manager tab. Existing .txt templates remain readable;
            # new native/OCR templates are JSON and can be inspected or removed here.
            template_tab = tk.Frame(notebook)
            notebook.add(template_tab, text="Template Manager")
            template_controls = tk.Frame(template_tab)
            template_controls.pack(side=tk.TOP, fill=tk.X, pady=2)
            tk.Button(template_controls, text="Refresh", command=self.refresh_template_manager).pack(side=tk.LEFT, padx=2)
            tk.Button(template_controls, text="Open Selected", command=self.open_template_item).pack(side=tk.LEFT, padx=2)
            tk.Button(template_controls, text="Delete Selected", command=self.delete_template_item).pack(side=tk.LEFT, padx=2)
            self.template_tree = ttk.Treeview(
                template_tab, columns=("Type", "Vendor", "File"), show="headings")
            self.template_tree.heading("Type", text="Type")
            self.template_tree.heading("Vendor", text="Vendor")
            self.template_tree.heading("File", text="File")
            self.template_tree.column("Type", width=100, anchor="center")
            self.template_tree.column("Vendor", width=220, anchor="w")
            self.template_tree.column("File", width=600, anchor="w")
            self.template_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            template_scroll = ttk.Scrollbar(template_tab, orient="vertical", command=self.template_tree.yview)
            template_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.template_tree.configure(yscrollcommand=template_scroll.set)
            self.template_tree.bind("<Double-1>", lambda event: self.open_template_item())
            self.template_items = {}
            self.refresh_template_manager()

            # About tab
            about_tab = tk.Frame(notebook)
            notebook.add(about_tab, text="About")
            about_text = ("Pewter Email Processor v1.0\n")
            tk.Message(
                about_tab,
                text=about_text,
                width=500,
                justify=tk.LEFT,
                font=("Courier", 11),
            ).pack(padx=10, pady=10)

            # Days without crashing counter with reset button, saves date in config
            counter_frame = tk.Frame(self.root) 
            counter_frame.pack(side=tk.BOTTOM, padx=1)
            self.days_without_crashing = tk.StringVar()
            self.load_crash_counter()
            self.update_crash_counter_label()
            tk.Label(counter_frame, textvariable=self.days_without_crashing).pack(side=tk.LEFT)
            tk.Button(counter_frame, text="↻", command=self.reset_crash_counter, width=2).pack(side=tk.LEFT)

            # GUI STYLES
            self.log_text_widget.tag_configure("red", background="#FFCCCC")
            self.log_text_widget.tag_configure("orange", background="#FFB434")
            self.log_text_widget.tag_configure("yellow", background="#FAFA33")
            self.log_text_widget.tag_configure("lgreen", background="#CCFFCC")
            self.log_text_widget.tag_configure("green", background="#39FF12")
            self.log_text_widget.tag_configure("dgreen", background="#00994d")
            self.log_text_widget.tag_configure("blue", background="#00FFFF")
            self.log_text_widget.tag_configure("purple", background="#CCCCFF")
            self.log_text_widget.tag_configure("gray", background="#DEDDDD")
            self.log_text_widget.tag_configure("no_new_emails", background="#DEDDDD")
            self.log_text_widget.tag_configure("label_error", background="#FFB434")
            self.log_text_widget.tag_configure("default", borderwidth=0.5, relief="solid", lmargin1=10, offset=8)  # applied to all messages
            
            # Redirect stdout and stderr to console
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

    def main(self):
        if self.TESTING:
            self.TEMPLATE_FOLDER = config["TEST_TEMPLATE_FOLDER"]
            self.OCR_TEMPLATE_FOLDER = config.get("TEST_OCR_TEMPLATE_FOLDER", config["OCR_TEMPLATE_FOLDER"])
            self.INVOICE_FOLDER = config["TEST_INVOICE_FOLDER"]
            self.log("Testing mode enabled", tag="yellow")
        else:
            self.TEMPLATE_FOLDER = config["TEMPLATE_FOLDER"]
            self.OCR_TEMPLATE_FOLDER = config["OCR_TEMPLATE_FOLDER"]
            self.INVOICE_FOLDER = config["INVOICE_FOLDER"]

        for folder in (self.TEMPLATE_FOLDER, self.OCR_TEMPLATE_FOLDER, self.INVOICE_FOLDER):
            if folder:
                os.makedirs(folder, exist_ok=True)
        with open(config["LOG_FILE"], "a", encoding="utf-8") as file:
            file.write("\n\n")
        self.log("Connecting...", tag="dgreen")
        self.processor_running = True
        self.ui(self.button_startup)

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
        self.pause_button.config(text="Pause", command=self.pause_processing, state=tk.NORMAL)
        self.pause_event.clear()
        self.logout_button.config(state=tk.NORMAL)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.testing_button.config(state=tk.DISABLED)
        self.away_mode_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.NORMAL)

    def button_logout(self):  
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.logout_button.config(state=tk.DISABLED)
        self.errors_button.config(state=tk.DISABLED)
        self.print_errors_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)

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
            if log:
                self.log(f"--- Connected to {self.username} --- {self.current_time} {self.current_date}", tag="dgreen")
            return imap
        except Exception as e:
            if primary:
                self.connected = False
            if log:
                self.log(f"Unable to connect to {self.username}: {e}", tag="red", send_email=True)
            raise

    def disconnect(self, imap, log=True):
        if imap is None:
            self.connected = False
            return
        try:
            imap.logout()
        except Exception as e:
            if log:
                self.log(f"An error occurred while disconnecting: {e}", tag="orange")
        finally:
            if imap is getattr(self, "imap", None):
                self.connected = False
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
            f"Invoice identity collision: {os.path.basename(desired_path)} already exists with different content; "
            f"saving as {os.path.basename(conflict)}.", tag="orange", send_email=True,
        )
        return conflict, False

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
            return
        
        self.archive.insert(
            "", "end",
            iid=new_id,
            values=(values[0], values[1], values[2], values[3], values[4], values[5], values[6]),
            tags=("default",))
        self.save_archive((new_id, values[0], values[1], values[2], values[3], values[4], values[5], values[6]))  # Save to archive
        self.inbox.delete(_id)  # Remove item from inbox
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
        print(values)
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
                            self.ui(self.errors_button.config, bg="#FBFF2C")
                        elif label == "Need_Print":
                            self.ui(self.print_errors_button.config, bg="#FBFF2C")
            except (imaplib.IMAP4.abort, socket.error, TimeoutError):
                raise

            except Exception as e:
                self.log(
                    f"An error occurred while checking the label: {str(e)}\n"
                    f"{traceback.format_exc()}",
                    tag="red",
                    send_email=True,
                )

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
                self.errors_button.config(bg="#F0F0F0")  # Reset button color
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
                self.print_errors_button.config(bg="#F0F0F0")  # Reset button color
        except Exception as e:
            self.log(f"Error resolving unprinted invoices: {str(e)}", tag="red", send_email=True)

    def pause_processing(self):  # Pauses processing
        self.log("Processing paused.", tag="yellow")
        self.pause_button.config(text="Resume", command=self.resume_processing)
        self.errors_button.config(state=tk.DISABLED)
        self.print_errors_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)
        self.pause_event.set()

    def resume_processing(self):  # Resumes processing
        self.log("Processing resumed.", tag="yellow")
        self.pause_button.config(text="Pause", command=self.pause_processing)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.test_inbox_button.config(state=tk.NORMAL)
        self.pause_event.clear()

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
        if self.email_executor is not None:
            self.email_executor.shutdown(wait=False, cancel_futures=False)
            self.email_executor = None

        if reconnect:
            # wait a few seconds then reconnect
            time.sleep(5)
            self.main()  # restart main loop

    def toggle_testing(self):  # Toggles testing mode
        if self.TESTING:
            self.TESTING = False
            self.testing_button.config(bg="#FFCCCC")
        else:
            self.TESTING = True
            self.testing_button.config(bg="#CCFFCC")

    def toggle_away_mode(self):  # Toggles away mode, always prints an saves invoice
        if self.AWAY_MODE:
            self.AWAY_MODE = False
            self.away_mode_button.config(bg="#FFCCCC")
        else:
            self.AWAY_MODE = True
            self.away_mode_button.config(bg="#CCFFCC")

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