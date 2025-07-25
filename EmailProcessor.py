import Rectangulator
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import threading
import importlib
import traceback
import win32print
import win32gui
import win32api
import imaplib
import smtplib
import sqlite3
import config
import email
import queue
import time
import sys
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class RedirectText:

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


class EmailProcessor:

    # CONSTANTS
    CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.py")
    ICON_PATH = os.path.join(os.path.dirname(__file__), "Hotpot.ico")
    TEMPLATE_FOLDER = config.TEMPLATE_FOLDER
    INVOICE_FOLDER = config.INVOICE_FOLDER
    ARCHIVE_DB = os.path.join(os.path.dirname(__file__), "archive.db")

    def __init__(self, username, password):
        try:
            # VARIABLES
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

            # GUI
            self.root = tk.Tk()
            self.root.iconbitmap(self.ICON_PATH)
            self.root.geometry("1400x700")
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
            right_frame = tk.Frame(program_tab)
            right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
            left_frame = tk.Frame(program_tab)
            left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

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
                height=30,
                width=140,
                spacing1=4,
                padx=0,
                pady=0,
            )  
            self.log_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.configure(command=self.log_text_widget.yview)


            # Plot canvas for rectangulator
            self.figure = Figure(figsize=(6.75, 6), dpi=100)
            self.ax = self.figure.add_subplot(111)
            self.ax.axis("off")
            self.canvas = FigureCanvasTkAgg(self.figure,
                                            master=right_frame)
            self.canvas.get_tk_widget().pack(side=tk.TOP,
                                             fill=tk.BOTH,
                                             expand=True)

            self.gui_queue = queue.PriorityQueue()  # queue for all gui tasks
            self.gui_busy = False
            self.rectangulator_handler = Rectangulator.RectangulatorHandler(
                self, self.figure, self.ax)


            # Archive tab
            archive_tab = tk.Frame(notebook)
            notebook.add(archive_tab, text="Archive")
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
            self.archive.pack(fill=tk.BOTH, expand=True)

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
            vars = {
                "APC_USER": tk.StringVar(value=config.APC_USER),
                "APC_PASS": tk.StringVar(value=config.APC_PASS),
                "LOG_FILE": tk.StringVar(value=config.LOG_FILE),
                "INVOICE_FOLDER": tk.StringVar(value=config.INVOICE_FOLDER),
                "TEMPLATE_FOLDER": tk.StringVar(value=config.TEMPLATE_FOLDER),
                "TEST_INVOICE_FOLDER": tk.StringVar(value=config.TEST_INVOICE_FOLDER),
                "TEST_TEMPLATE_FOLDER": tk.StringVar(value=config.TEST_TEMPLATE_FOLDER),
                "INBOX_CYCLE_TIME": tk.IntVar(value=config.INBOX_CYCLE_TIME),
                "RECONNECT_TIME": tk.IntVar(value=config.RECONNECT_TIME),
                "RECEIVER_EMAIL": tk.StringVar(value=config.RECEIVER_EMAIL),
            }
            for i, (key, var) in enumerate(vars.items()):
                label = tk.Label(settings_tab, text=key + ":")
                label.grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
                entry = tk.Entry(settings_tab, textvariable=var, width=100)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky=tk.W)

            def save_settings():
                # Saves settings to config.py
                try:
                    lines = []
                    with open(self.CONFIG_PATH, "r") as f:
                        for line in f:
                            for key, var in vars.items():
                                if line.startswith(key):
                                    line = f"{key} = {repr(var.get())}\n"
                            lines.append(line)
                                    
                    with open(self.CONFIG_PATH, "w") as f:
                        f.writelines(lines)
                    messagebox.showinfo("Settings", "Settings saved successfully.")
                    importlib.reload(config)  # Reload config module to apply changes
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
            save_button = tk.Button(settings_tab, 
                                    text="Save Settings",
                                    command=save_settings)
            save_button.grid(row=len(vars), column=0, columnspan=2, pady=10)

            # About tab
            about_tab = tk.Frame(notebook)
            notebook.add(about_tab, text="About")
            about_text = (
                "Pewter Email Processor v1.0\n"
            )
            tk.Message(
                about_tab,
                text=about_text,
                width=500,
                justify=tk.LEFT,
                font=("Courier", 11),
            ).pack(padx=10, pady=10)

            # Days without crashing counter with reset button, saves date in config
            counter_frame = tk.Frame(self.root) 
            counter_frame.pack(side=tk.RIGHT, padx=1)
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

            self.root.protocol(
                "WM_DELETE_WINDOW",
                self.on_program_exit)  # runs exit protocol on window close
            self.root.after(1000, self.process_queue)  # starts queue processing
            self.root.mainloop()
        except Exception as e:
            print(f"-An error occurred while initializing the EmailProcessor: {str(e)}")

    def main(self):  # Runs when start button is pressed
        if self.TESTING:  # set appropriate folders for testing
            self.TEMPLATE_FOLDER = config.TEST_TEMPLATE_FOLDER
            self.INVOICE_FOLDER = config.TEST_INVOICE_FOLDER
            self.log("Testing mode enabled", tag="yellow")
        else:
            self.TEMPLATE_FOLDER = config.TEMPLATE_FOLDER
            self.INVOICE_FOLDER = config.INVOICE_FOLDER

        with open(config.LOG_FILE, "a") as file:  # open log file
            file.write("\n\n")
        self.log("Connecting...", tag="dgreen")
        self.root.update()
        self.processor_running = True

        # Enable and disable buttons
        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(text="Pause",
                                 command=self.pause_processing,
                                 state=tk.NORMAL)
        self.pause_event.clear()
        self.logout_button.config(state=tk.NORMAL)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.testing_button.config(state=tk.DISABLED)
        self.away_mode_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.NORMAL)

        # Login primary connnection to gmail
        self.imap = self.connect()
        self.imap_lock = threading.RLock()  # lock for imap connection
        if self.imap:
            self.processor_thread = threading.Thread(target=self.search_inbox)
            self.processor_thread.start()

    def safe_imap(self, func, *args, **kwargs):
        """Safely calls an imap function with error handling."""
        try:
            with self.imap_lock:  # ensure thread-safe access to imap
                return func(*args, **kwargs)
        except imaplib.IMAP4.abort as e:
            self.log(f"Socket error: {str(e)} -- {self.current_time} {self.current_date}", tag="red", send_email=False)
            self.logout()
            raise
        except Exception as e:
            self.log(f"An error occurred: {str(e)}", tag="red", send_email=True)
            raise 

    def connect(self, log=True):  # Connects email, returns imap object
        user = f"{self.username}{config.ADDRESS}"
        try:
            imap = imaplib.IMAP4_SSL(config.IMAP_SERVER)
            imap.socket().settimeout(10)
            imap.login(user, self.password)
            self.connected = True
            if log:
                self.log(
                    f"--- Connected to {self.username} --- {self.current_time} {self.current_date}",
                    tag="dgreen",)
            return imap
        except imaplib.IMAP4.error as e:
            if log:
                self.log(
                    f"Unable to connect to {self.username}: {str(e)}",
                    tag="red",
                    send_email=True,)

    def disconnect(self, imap, log=True):  # Disconnects email
        try:
            self.safe_imap(imap.logout)  # safely logout
            self.connected = False
            if log:
                self.log(
                    f"--- Disconnected from {self.username} --- {self.current_time} {self.current_date}",
                    tag="red",)
        except Exception as e:
            if log:
                self.log(
                    f"An error occurred while disconnecting: {str(e)}",
                    tag="red",
                    send_email=True,)

    """ 
    Search inbox loops through emails, creates thread for each email to process it by running process_email.
    Process email calls handle_attachments which calls add_to_queue, which adds the pdf to the queue which is processed by process_queue on a separate thread.
    The queue is processed by handle_rectangulator which runs the rectangulator and prints the invoice if needed.
    """

    def search_inbox(self):  # Main Loop, searches inbox for new emails, creates thread to handle each
        try:
            cycle_count = 0  # cycle count for reconnecting every hour
            while self.processor_running:
                if not self.pause_event.is_set() and self.connected:
                    # Search for all emails in the inbox
                    self.safe_imap(self.imap.select, "INBOX")  # select inbox
                    _, emails = self.safe_imap(self.imap.uid, "search", None, "UNSEEN")
                    uids = [uid.decode() for uid in emails[0].split() if uid]  # Decode bytes to str and filter out empty uids
                    new_uids = []

                    # check if emails have already been processed
                    with self.current_emails_lock:
                        for uid in uids:
                            if uid not in self.current_emails:
                                new_uids.append(uid)
                                self.current_emails.add(uid)
                                self.safe_imap(self.imap.uid, "STORE", uid, "-FLAGS", "(\Seen)")  # mark as unseen

                    # Check if no new mail
                    if not new_uids:
                        self.log(
                            f"No new emails - {self.current_time} {self.current_date}",
                            tag="no_new_emails",
                            write=False)
                        self.check_labels(["Need_Print", "Errors"], self.imap)
                        self.pause_event.wait(timeout=config.INBOX_CYCLE_TIME)  # pause until next cycle
                    else:
                        for uid in new_uids:
                            threading.Thread(target=self.process_email,
                                             args=(uid, ),
                                             daemon=True).start()
                        self.root.after(0, self.flash_taskbar)  # flash taskbar if new email

                    cycle_count += 1
                    # Reconnect every hour
                    if cycle_count % config.RECONNECT_CYCLE_COUNT == 0:
                        self.reconnect()
                        cycle_count = 0

            # Disconnect when the program is closed
            self.disconnect(self.imap)
            if self.logging_out:
                self.logging_out = False
                self.start_button.config(state=tk.NORMAL)
                self.testing_button.config(state=tk.NORMAL)
                self.away_mode_button.config(state=tk.NORMAL)
        except imaplib.IMAP4.abort as e:
            self.log(f"Search Socket error: {str(e)} -- {self.current_time} {self.current_date}", tag="red", send_email=False)
            self.logout()
        except Exception as e:
            self.log(
                f"An error occurred while searching the inbox: {str(e)} \n{traceback.format_exc()}",
                tag="red",
                send_email=True,)
            self.restart_processing()

    def process_email(self, mail):  # Handles each email in own thread created by search_inbox
        subject = ""
        imap = self.connect(log=False)
        try:
            # Fetch email
            msg = self.get_msg(mail, "INBOX", imap)
            if msg is None:
                self.log(
                    f"Email not found (process_email): {mail}",
                    tag="red",
                    send_email=True)
                return
            subject = msg["Subject"]
            sender_email = email.utils.parseaddr(msg["From"])[1]

            # Check if sender is trusted
            if not sender_email.endswith(config.TRUSTED_ADDRESS):
                self.move_email(mail, "Not_Invoices", "INBOX", imap)
                return

            # Check for attachments
            has_attachment = any(
                part.get_content_disposition() == "attachment"
                and part.get_filename().lower().endswith(".pdf")
                for part in msg.walk())
            if not has_attachment:
                pass
            else:
                with self.current_emails_lock:
                    self.current_emails.add(mail)
                self.handle_attachments(mail, imap, subject)
        except Exception as e:
            self.log(
                f"An error occurred while processing an email: {str(e)} \n{traceback.format_exc()}",
                tag="red",
                send_email=True)
            self.move_email(mail, "Errors", "INBOX", imap)
            return
        finally:
            try:
                print(f"-Disconncting imap for email {mail} - {subject}")
                imap.logout()
            except:
                pass

    def handle_attachments(self, mail, imap, subject):  # Iterate over email parts and find pdf, ran by process_email
        msg = self.get_msg(mail, "INBOX", imap)
        if msg is None:
            self.log(
                f"Email not found (handle_attachments): {mail}",
                tag="red",
                send_email=True)
            return
        filenames = []

        for part in msg.walk():
            if (part.get_filename() not in filenames
                    and part.get_content_disposition() is not None
                    and part.get_filename() is not None
                    and part.get_filename().lower().endswith(".pdf")):
                filenames.append(part.get_filename())
                if subject == "Test":
                    self.add_to_queue(mail, part, subject, testing="test")
                elif self.TESTING:
                    self.add_to_queue(mail, part, subject, testing=True)
                else:
                    self.add_to_queue(mail, part, subject)
        self.remaining_pdfs[mail] = len(filenames)
        self.log(f"Processing {len(filenames)} PDF(s) for email '{subject}'",
                 tag="blue")

    def add_to_queue(self, mail, part, subject, testing=False):  # Adds invoices to queue, ran by handle_attachments
        try:
            # Get fllename and attachment
            filename = part.get_filename()
            if testing == "test":  # when testing inbox
                filename = f"Test_{filename}"
            filepath = os.path.join(self.INVOICE_FOLDER, filename)
            attachment = part.get_payload(decode=True)

            # Check if file already exists
            if os.path.exists(filepath):
                filename = f"{filename[:-4]}_{str(int(time.time()))}.pdf"
                filepath = os.path.join(self.INVOICE_FOLDER, filename)

            # Download invoice PDF
            with open(filepath, "wb") as file:
                file.write(attachment)

            self.gui_queue.put((1, "NEW", mail, subject, filename, filepath))
            self.gui_queue.put((2, "RECTANGULATE", mail, filename, filepath, testing))  # add to queue for rectangulator
        except Exception as e:
            self.log(
                f"An error occurred while processing the queue: {str(e)}",
                tag="red",
                send_email=True)

    def process_queue(self):  # Processes queue of emails
        try:
            # Check if any gui tasks are in the queue
            if not self.gui_queue.empty() and not self.gui_busy:
                task = self.gui_queue.get()[1:]  # Unpack the task, ignore priority
                task_type = task[0]
                if task_type == "NEW":
                    mail, subject, filename, filepath = task[1:5]
                    self.inbox.insert(
                        "", "end",
                        iid=f"{mail}_{filename}",
                        values=(subject, self.current_date, filename, "No", "No", "", filepath), # Invoice, Saved, Printed, Errors
                        tags=("pending",))
                    
                if task_type == "STATUS":
                    id, status = task[1:3]
                    item = self.inbox.item(id)
                    values = list(item["values"])

                    if status == "saved":
                        filename, filepath = task[3:5]
                        values[2] = filename
                        values[6] = filepath
                        values[3] = "Yes"
                        self.inbox.item(id, tags=("finished",))
                    elif status == "printed":
                        values[4] = "Yes"
                    elif status.startswith("Error"):
                        values[5] = status
                        self.inbox.item(id, tags=("error",))
                    self.inbox.item(id, values=values)

                if task_type == "REMOVE":
                    id = task[1]
                    if id in self.inbox.get_children():
                        self.inbox.delete(id)

                if task_type == "RECTANGULATE":
                    self.gui_busy = True
                    # Unpack the task
                    mail, filename, filepath, testing = task[1:6]
                    # run rectangulator in the main thread
                    self.root.after(0,
                        lambda: self.handle_rectangulator(
                            mail, filename, filepath, testing),)  
        finally:
            self.root.after(100, self.process_queue)  # Schedule the next check of the queue

    def handle_rectangulator(self, mail, filename, filepath, testing):  # Handles rectangulator
        try:
            inbox_item_id = f"{mail}_{filename}"
            return_list = self.rectangulator_handler.rectangulate(
                filename, filepath, self, self.TEMPLATE_FOLDER, testing)
            print(f"-rectangulator returned: {return_list}")

            # Check if Rectangulator fails
            if return_list == [] or return_list[0] == None:
                self.move_email(mail, "Errors", "INBOX", self.imap)
                os.remove(filepath)
                self.log(
                    f"Failed to download '{filename}', moved to Error label, not printed",
                    tag="red",
                    send_email=True)
                self.gui_queue.put((0, "STATUS", inbox_item_id, "Error Rectangulating"))
                return

            # Check if test email
            if return_list[0] == "test_email":       
                self.log(
                    f"Test complete",
                    tag="purple")
                self.move_email(mail, "Test_Email", "INBOX", self.imap)
                os.remove(filepath)
                self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", "Test Email", "None"))
                return

            # Check if email is already processed
            if mail in self.remaining_pdfs:
                self.remaining_pdfs[mail] -= 1
                if self.remaining_pdfs[mail] == 0:
                    del self.remaining_pdfs[mail]
                    self.move_email(mail, "Invoices", "INBOX", self.imap)

            new_filepath, should_print = return_list
            # Check if template used
            if should_print == "template":
                self.log(f"Created new invoice file {os.path.basename(new_filepath)} -- {self.current_date} {self.current_time}", tag="lgreen")
                if not testing:
                    self.print_invoice(new_filepath, inbox_item_id)
                self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", os.path.basename(new_filepath), new_filepath))
                return

            # Check if not invoice
            if new_filepath == "not_invoice":
                new_filepath, should_print, should_save = should_print    
                if testing == True:
                    should_print = False
                if should_print:
                    self.print_invoice(filepath, inbox_item_id)
                if not should_save:
                    self.log(f"{os.path.basename(new_filepath)} marked not an invoice and not saved -- {self.current_time} {self.current_date}", tag="purple")
                    os.remove(filepath)
                    self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", "Not Invoice", "None"))     
                    return
                self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", "Not Invoice", new_filepath))
                self.log(f"{os.path.basename(new_filepath)} marked not an invoice and saved -- {self.current_time} {self.current_date}", tag="purple")
            
            if testing == True:
                should_print = False

            # Check if invoice has already been processed
            if os.path.exists(new_filepath):
                old_filepath = new_filepath
                self.log(
                    f"New invoice file already exists at {os.path.basename(old_filepath)} -- {self.current_time} {self.current_date}",
                    tag="orange")
                new_filepath = f"{old_filepath[:-4]}_{str(int(time.time()))}.pdf"

            # Save invoice
            os.rename(filepath, new_filepath)
            self.gui_queue.put((0, "STATUS", inbox_item_id, "saved", os.path.basename(new_filepath), new_filepath))
            self.log(
                f"Created new invoice file {os.path.basename(new_filepath)} -- {self.current_date} {self.current_time}",
                tag="lgreen")
            if should_print:
                self.print_invoice(new_filepath, inbox_item_id)
        finally:
            self.gui_busy = False

    def remove_inbox_item(self, event):  # Removes inbox item on double click and adds to archive
        # Only if item has finished tag
        item = self.inbox.selection()  # Get selected item
        if not item or "finished" not in self.inbox.item(item[0], "tags"):
            return
        
        id = item[0]
        values = self.inbox.item(id, "values")  # Get values of selected item
        new_id = f"{id}_{int(time.time())}" # create new unique id for archive item

        if values[2] == "Test Email":  # If item is test email, don't archive
            self.inbox.delete(id)
            return
        
        self.archive.insert(
            "", "end",
            iid=new_id,
            values=(values[0], values[1], values[2], values[3], values[4], values[5], values[6]),
            tags=("default",))
        self.save_archive((new_id, values[0], values[1], values[2], values[3], values[4], values[5], values[6]))  # Save to archive
        self.inbox.delete(id)  # Remove item from inbox
        self.log(f"Archived inbox item {id} as {new_id}.", tag="blue")

    def remove_archive_item(self, event):  # Removes archive item on double right click
        item = self.archive.selection()  # Get selected item
        if not item:
            return
        
        id = item[0]
        self.archive.delete(id)  # Remove item from archive tree
        self.db.execute("DELETE FROM archive WHERE id = ?", (id,))  # Remove from database
        self.db.commit() 
        self.log(f"Removed archive item {id} from archive.", tag="blue")  # Log removal

    def print_archive_item(self, event): # Prints archive item on middle click
        item = self.archive.selection()
        if not item:
            return
        
        id = item[0]
        filepath = self.archive.item(id, "values")[-1]
        if not filepath or not os.path.exists(filepath):
            self.log(f"File not found for printing archive item {id}, {filepath}", tag="red")
            return
        try:
            self.print_invoice(filepath)
        except Exception as e:
            self.log(f"Error printing archive item {id}: {str(e)}", tag="red")


    def archive_all(self):  # Archives all inbox items
        for item in self.inbox.get_children(): # select item and call remove
            self.inbox.selection_set(item)  # Select item
            self.remove_inbox_item(None)  # Call remove inbox item method

    def save_archive(self, record):
        self.db.execute("INSERT OR REPLACE INTO archive VALUES (?, ?, ?, ?, ?, ?, ?, ?)", record)
        self.db.commit()

    def load_archive(self):  # Loads archive from database
        for row in self.db.execute("SELECT * FROM archive ORDER BY datestamp DESC"):
            self.archive.insert("", "end", iid=row[0], values=row[1:])

    def open_archive_item(self, event):  # Opens archive item on double click
        item = self.archive.selection()
        if not item:
            return
        id = item[0]
        filepath = self.archive.item(id, "values")[-1]  # Get filepath from selected item
        if not filepath or not os.path.exists(filepath):
            self.log(f"File not found for archive item {id}: {filepath}",
                     tag="red",
                     send_email=True)
            return
        try:
            os.startfile(filepath)  # Open the file
        except Exception as e:
            self.log(f"Error opening file {filepath}: {str(e)}",
                     tag="red",)
            
    def move_email(self, mail, label, og_label, imap):  # Moves email to label
        subject = "Unknown"
        try:
            with self.imap_lock:
                # Get msg and subject if possible
                msg = self.get_msg(mail, og_label, imap)
                if msg:
                    subject = msg["Subject"]

                # Make a copy of the email in the specified label
                imap.select(og_label)
                success = imap.uid("COPY", mail, label)
                if success[0] != "OK":
                    self.log(
                        f"Error copying email '{subject}': {success[1]}",
                        tag="red",
                        send_email=True)
                    return False
                # Mark as unseen
                imap.uid("STORE", mail, "-FLAGS", "(\Seen)")

                # Mark the original email as deleted
                imap.uid("STORE", mail, "+FLAGS", "(\Deleted)")
                imap.expunge()
                self.log(f"Moved email '{subject}' from {og_label} to {label}.",
                        tag="blue")
        except Exception as e:
            self.log(f"Transfer failed for '{subject}': {str(e)} \n{traceback.format_exc()}",
                     tag="red",
                     send_email=True)

    def send_email(self, body):  # Sends email to me
        sender_email = f"{self.username}{config.ADDRESS}"
        try:
            if self.TESTING:
                return

            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["Subject"] = "Alert"
            message["From"] = sender_email
            message["To"] = config.RECEIVER_EMAIL
            message.attach(MIMEText(body, "plain"))

            # Send the email using SMTP
            with smtplib.SMTP(config.SMTP_SERVER, 587) as server:
                server.starttls()
                server.login(sender_email, self.password)
                server.sendmail(sender_email, config.RECEIVER_EMAIL,
                                message.as_string())
        except Exception as e:
            self.log(f"Error sending email - {str(e)}", tag="red")

    def get_subject(self, mail, label, imap):  # Get the message from the specified email
        try:
            with self.imap_lock:
                imap.select(label)
                result, data = imap.uid("FETCH", mail, "(RFC822)")
                if result != "OK" or not data or not data[0]:
                    self.log(f"Error fetching email: {mail}",
                            tag="red",
                            send_email=True)
                    return None
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                return msg["Subject"]
        except Exception as e:
            self.log(f"Error getting subject: {str(e)}",
                     tag="red",
                     send_email=True)
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

    def print_invoice(self, filepath, inbox_item_id=None):  # Printer
        try:
            # Get default printer and print
            p = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", filepath, None, ".", 0)
            self.log(f"Printed {os.path.basename(filepath)}.", tag="lgreen")
            if inbox_item_id:
                self.gui_queue.put((0, "STATUS", inbox_item_id, "printed"))
            return True
        except Exception as e:
            self.log(f"Printing failed for {filepath}: {str(e)}",
                     tag="red",
                     send_email=True)
            if inbox_item_id:
                self.gui_queue.put((0, "STATUS", inbox_item_id, "Error Printing"))
            return False

    def log(self, *args, tag=None, send_email=False, write=True):  # Logs to text box and log file
        try:
            if self.window_closed:  # check if window is still open
                return
            message = " ".join([str(arg) for arg in args])  # convert args to string

            # Get rid of no_new_emails and label_error messages
            if tag == "no_new_emails" or tag == "label_error":
                self.remove_messages(message)
            else:
                print(f"-{message}")

            # Insert the new message to the text widget
            self.log_text_widget.insert(tk.END, message + "\n", (tag, "default"))
            # If the bottom quarter of the text widget is visible, autoscroll
            if self.log_text_widget.yview()[1] > 0.75:
                self.log_text_widget.yview_moveto(1)

            # Send email for errors
            if tag == "red" and send_email:
                self.send_email(message)

            # Write to the log file
            if write:
                with open(config.LOG_FILE, "a") as file:
                    file.write(message + "\n")
        except Exception as e:
            print(f"-Error logging: {str(e)}")

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
                    if len(labels) == 1:
                        return email_ids
                    elif len(email_ids) > 0:
                        self.log(
                            f"{len(email_ids)} emails in {label} - {self.current_time} {self.current_date}",
                            tag="label_error",
                            write=False)
            except Exception as e:
                self.log(
                    f"An error occurred while checking the label: {str(e)}",
                    tag="red",
                    send_email=True)
                raise imaplib.IMAP4.abort

    def get_msg(self, mail, label, imap):  # Gets email message
        try:
            with self.imap_lock:
                imap.select(label)
                result, data = imap.uid("FETCH", mail, "(RFC822)")
                if result != "OK" or not data or not data[0]:
                    self.log(f"Error fetching email: {mail}",
                            tag="red",
                            send_email=True)
                    return None
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                return msg
        except Exception as e:
            self.log(f"Error getting message: {str(e)}",
                     tag="red",
                     send_email=True)
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
        except Exception as e:
            self.log(f"Error resolving errors: {str(e)}",
                     tag="red",
                     send_email=True)

    def resolve_prints(self):  # Moves unprinted invoices back to inbox
        try:
            self.log(f"Attempting to resolve unprinted invoices.",
                     tag="yellow")
            # Get emails in Need_Print label
            email_ids = self.check_labels(["Need_Print"], self.imap)

            if len(email_ids) == 0:
                self.log(f"No unprinted invoices to resolve.", tag="yellow")
                return

            # Move emails back to inbox
            for email_id in email_ids:
                self.move_email(email_id, "INBOX", "Need_Print", self.imap)
        except Exception as e:
            self.log(
                f"Error resolving unprinted invoices: {str(e)}",
                tag="red",
                send_email=True)

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
        self.log(f"Restarting...", tag="yellow")
        self.disconnect(self.imap)
        if self.processor_thread:
            self.pause_event.set()  # pause processing
            self.processor_thread.join()  # wait for thread to finish
        self.processor_running = False
        self.main()

    def logout(self, reconnect=False):  # Logs out
        self.log("Logging out...", tag="yellow")
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.logout_button.config(state=tk.DISABLED)
        self.errors_button.config(state=tk.DISABLED)
        self.print_errors_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)
        self.pause_event.set()
        self.processor_running = False
        self.logging_out = True
        self.current_emails.clear()  # clear current emails

        if reconnect:
            # wait a few seconds then reconnect
            time.sleep(5)
            self.reconnect()

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

    def test_rectangulator(self):  # Opens rectangulator with test invoice
        self.log("Testing Rectangulator...", tag="yellow")
        return_list = []
        return_list = self.rectangulator_handler.rectangulate(
            "Testing Rectangulator", config.TEST_INVOICE, self, config.TEST_TEMPLATE_FOLDER, True)
        if return_list != []:
            new_filepath, should_print = return_list
            self.log(
                f"Test complete - new_filepath: {new_filepath}, should_print: {should_print}",
                tag="purple")
        self.log("Testing complete.", tag="yellow")

    def test_inbox(self):  # Sends test email to inbox, won't be printed or downloaded
        self.log("Sending test email to inbox", tag="yellow")
        mail = self.get_email("Test_Email", self.imap)
        if mail is None:
            self.log("Test email missing", tag="orange")
            return
        self.move_email(mail, "INBOX", "Test_Email", self.imap)

    def reconnect(self):  # Reconnects to email
        self.disconnect(self.imap, log=False)
        self.imap = self.connect(log=False)
        self.log(
            f"Reconnected to {self.username} -- {self.current_time} {self.current_date}",
            tag="green")
        self.update_crash_counter_label()

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

        # Destroys tkinter window
        self.root.destroy()

    def flash_taskbar(self):  # Flash icon in taskbar
        # Code from stack overflow
        hwnd_int = int(self.root.frame(), base=16)
        win32gui.FlashWindow(hwnd_int, 0)

    def load_crash_counter(self): # Sets date variable from config
        try:
            last_crash_date = config.LAST_CRASH_DATE.strip()
            self.last_crash_date = last_crash_date if last_crash_date else str(datetime.now().strftime("%Y-%m-%d"))
        except Exception as e:
            print(f"-No crash date found {str(e)}")
            self.last_crash_date = str(datetime.now().strftime("%Y-%m-%d"))
            self.save_crash_counter()

    def save_crash_counter(self): # Overwrites date in config
        try:
            lines = []
            with open(self.CONFIG_PATH, "r") as f:
                for line in f:
                    if line.startswith("LAST_CRASH_DATE"):
                        line = f"LAST_CRASH_DATE = '{self.last_crash_date}'"
                    lines.append(line)
            with open(self.CONFIG_PATH, "w") as f:
                f.writelines(lines)
            importlib.reload(config)  # Reload config module to apply changes
        except Exception as e:
            print(f"-what happened {str(e)}")

    def reset_crash_counter(self): # Reset date variable and updates label
        self.last_crash_date = str(datetime.now().strftime("%Y-%m-%d"))
        self.save_crash_counter()
        self.update_crash_counter_label()

    def update_crash_counter_label(self):
        self.days_without_crashing.set(f"Days without crashing: {(datetime.today() - datetime.strptime(self.last_crash_date, '%Y-%m-%d')).days}")

    @property
    def current_time(self):
        return time.strftime("%H:%M:%S", time.localtime())

    @property
    def current_date(self):
        return time.strftime("%m-%d-%Y", time.localtime())