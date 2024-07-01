import Rectangulator
import Loginulator
import threading
import traceback
import imaplib
import smtplib
import config
import email
import time
import os
import tkinter as tk
import win32gui
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class EmailProcessor:

    # CONSTANTS
    LOG_FILE = config.LOG_FILE
    ICON_PATH = os.path.join(os.path.dirname(__file__), "Hotpot.ico")
    TEMPLATE_FOLDER = config.TEMPLATE_FOLDER
    INVOICE_FOLDER = config.INVOICE_FOLDER
    IMAP_SERVER = config.IMAP_SERVER
    SMTP_SERVER = config.SMTP_SERVER
    RECIEVER_EMAIL = config.RECIEVER_EMAIL
    TRUSTED_ADDRESS = config.TRUSTED_ADDRESS
    ADDRESS = config.ADDRESS
    CYCLE_TIME = config.INBOX_CYCLE_TIME
    RECONNECT_CYCLE_COUNT = config.RECONNECT_CYCLE_COUNT
    TEST_INVOICE = config.TEST_INVOICE
    TEST_INVOICE_FOLDER = config.TEST_INVOICE_FOLDER
    TEST_TEMPLATE_FOLDER = config.TEST_TEMPLATE_FOLDER

    def __init__(self, username, password, rectangulator_handler):
        try:
            # GUI
            self.root = tk.Tk()
            self.root.iconbitmap(self.ICON_PATH)
            self.root.geometry("900x400")
            self.root.title(f"{username.upper()} Pewter")

            # VARIABLES
            self.username = username
            self.password = password
            self.rectangulator_handler = rectangulator_handler
            self.alert_window = None # used for pop ups
            self.window_closed = None
            self.processor_thread = None
            self.processor_running = False
            self.pause_event = threading.Event() # used for cycles
            self.connected = False
            self.logging_out = False
            self.TESTING = False # default to false

            # GUI BUTTONS
            self.button_frame = tk.Frame(self.root)
            self.button_frame.pack(side=tk.TOP)

            self.start_button = tk.Button(self.button_frame, text="Start", command=self.main) # start process button
            self.start_button.pack(side=tk.LEFT, padx=1)

            self.pause_button = tk.Button(self.button_frame, text="Pause", command=self.pause_processing, state=tk.DISABLED) # pause button
            self.pause_button.pack(side=tk.LEFT, padx=1)

            self.logout_button = tk.Button(self.button_frame, text="Logout", command=self.logout, state=tk.DISABLED) # logout button
            self.logout_button.pack(side=tk.LEFT, padx=1)

            self.errors_button = tk.Button(self.button_frame, text="Resolve Errors", command=self.resolve_errors, state=tk.DISABLED) # resolve errors button
            self.errors_button.pack(side=tk.LEFT, padx=1)

            self.print_errors_button = tk.Button(self.button_frame, text="Resolve Prints", command=self.resolve_prints, state=tk.DISABLED) # resolve unprinted invoices button
            self.print_errors_button.pack(side=tk.LEFT, padx=1)

            self.clear_button = tk.Button(self.button_frame, text="Clear", command=lambda: self.log_text_widget.delete("1.0", tk.END), state=tk.NORMAL) # clear button
            self.clear_button.pack(side=tk.LEFT, padx=1)

            self.testing_button = tk.Button(self.button_frame, text="Testing", command=self.toggle_testing, state=tk.NORMAL, bg="#FFCCCC", fg="black") # testing button
            self.testing_button.pack(side=tk.LEFT, padx=1)

            self.test_rectangulator_button = tk.Button(self.button_frame, text="Test Rectangulator", command=self.test_rectangulator, state=tk.NORMAL) # test rectangulator button
            self.test_rectangulator_button.pack(side=tk.LEFT, padx=1)

            self.test_inbox_button = tk.Button(self.button_frame, text="Test Inbox", command=self.test_inbox, state=tk.DISABLED) # test inbox button
            self.test_inbox_button.pack(side=tk.LEFT, padx=1)

            scrollbar = tk.Scrollbar(self.root)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            self.log_text_widget = tk.Text(self.root, yscrollcommand=scrollbar.set, height=30, width=140, spacing1=4, padx=0, pady=0) # text label
            self.log_text_widget.pack(side=tk.LEFT, fill=tk.BOTH)
            scrollbar.configure(command=self.log_text_widget.yview)

            # GUI STYLES
            self.log_text_widget.tag_configure("red", background="#FFCCCC")
            self.log_text_widget.tag_configure("yellow", background="yellow")
            self.log_text_widget.tag_configure("orange", background="#FFB434")	
            self.log_text_widget.tag_configure("lgreen", background="#CCFFCC") # light green
            self.log_text_widget.tag_configure("green", background="#39FF12") # green
            self.log_text_widget.tag_configure("dgreen", background="#00994d") # dark green
            self.log_text_widget.tag_configure("blue", background="#89CFF0")
            self.log_text_widget.tag_configure("purple", background="#E6E6FA")
            self.log_text_widget.tag_configure("gray", background="#DEDDDD")
            self.log_text_widget.tag_configure("no_new_emails", background="#DEDDDD") # gray
            self.log_text_widget.tag_configure("default", borderwidth=0.5, relief="solid", lmargin1=10, offset=8) # default

            self.root.protocol("WM_DELETE_WINDOW", self.on_program_exit) # runs exit protocol on window close
            self.root.mainloop()
        except Exception as e:
            print(f"An error occurred while initializing the EmailProcessor: {str(e)}")
            time.sleep(5) # try again after 5 seconds
            return self.__init__(username, password)
        

    def main(self): # Runs when start button is pressed 
        if self.TESTING:
            self.TEMPLATE_FOLDER = self.TEST_TEMPLATE_FOLDER
            self.INVOICE_FOLDER = self.TEST_INVOICE_FOLDER
            self.log("Testing mode enabled", tag="orange")

        self.log("Connecting...", tag="dgreen")
        self.root.update()
        self.processor_running = True

        # Enable and disable buttons
        self.start_button.config(state=tk.DISABLED) 
        self.pause_button.config(text="Pause", command=self.pause_processing, state=tk.NORMAL)
        self.pause_event.clear()
        self.logout_button.config(state=tk.NORMAL)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.testing_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.NORMAL)
        
        # Imap login
        self.imap = self.connect()
        if self.imap:
            self.processor_thread = threading.Thread(target=self.search_inbox)
            self.processor_thread.start()


    def connect(self, log=True): # Connects email, returns imap object
        user = f"{self.username}{self.ADDRESS}"
        try:
            imap = imaplib.IMAP4_SSL(self.IMAP_SERVER)
            imap.login(user, self.password)
            self.connected = True
            if log:
                self.log(f"--- Connected to {self.username} --- {self.current_time} {self.current_date}", tag="dgreen")
            return imap
        except imaplib.IMAP4_SSL.error as e:
            if log:
                self.log(f"Unable to connect to {self.username}: {str(e)}", tag="red", send_email=True)
            time.sleep(5)
            return self.connect(self.username, self.password, log=False) # try again after 5 seconds
        except RecursionError as e: 
            self.log(f"An error occurred while connecting: {str(e)}", tag="red", send_email=True)
            time.sleep(60) # wait a minutes
            return self.connect(self.username, self.password, log=False)
        

    def disconnect(self, log=True): # Disconnects email
        try:
            self.imap.logout()
            self.connected = False
            if log:
                self.log(f"--- Disconnected from {self.username} --- {self.current_time} {self.current_date}", tag="red")
        except Exception as e:
            if log:
                self.log(f"An error occurred while disconnecting: {str(e)}", tag="red", send_email=True)   
                self.disconnect(log=False) # try again after 5 seconds 
            else:
                # If disconnecting isn't working, were probably already disconnected
                self.log(f"Disconnecting isn't working: {str(e)}", tag="red")


    def search_inbox(self): # Main Loop, searches inbox for new emails
        try:
            cycle_count = 0
            while self.processor_running:
                if not self.pause_event.is_set() and self.connected:
                    # Search for all emails in the inbox
                    self.imap.select("inbox")
                    _, emails = self.imap.search(None, "ALL")

                    # Check if no new mail
                    if not emails[0]:
                        self.log(f"No new emails - {self.current_time} {self.current_date}", tag="no_new_emails")
                        self.check_labels(["Need_Print", "Need_Login", "Errors"])
                        self.pause_event.wait(timeout=self.CYCLE_TIME)  # pause until next cycle
                    else:
                        self.flash_taskbar()
                        self.process_email(emails[0].split()[0])

                    cycle_count += 1
                    # Reconnect every hour
                    if cycle_count == self.RECONNECT_CYCLE_COUNT:
                        self.reconnect()
                        cycle_count = 0
                        
            # Disconnect when the program is closed
            self.disconnect()
            if self.logging_out:
                self.logging_out = False
                self.start_button.config(state=tk.NORMAL)
                self.testing_button.config(state=tk.NORMAL)
        except imaplib.IMAP4.abort as e:
            self.log(f"Socket error: {str(e)}", tag="red", send_email=False)
            self.imap = self.connect(log=False)
            self.restart_processing()
        except Exception as e:  
            self.log(f"An error occurred while searching the inbox: {str(e)}", tag="red", send_email=True)
            self.restart_processing()


    def process_email(self, mail): # Handles each email
        subject = ""
        try: 
            # Fetch email
            msg = self.get_msg(mail, "inbox")
            subject = msg["Subject"]
            sender_email = email.utils.parseaddr(msg["From"])[1]

            # Check if sender is trusted
            if not sender_email.endswith(self.TRUSTED_ADDRESS):
                self.move_email(mail, "Not_Invoices", "inbox")
                return
            
            # Check for attachments
            has_attachment = any(part.get("content-disposition", "").startswith("attachment") for part in msg.walk() if msg.is_multipart())
            if not has_attachment:
                attachment_error = self.handle_login(mail)
            else:
                self.move_email(mail, "Queued", "inbox")
                self.handle_attachments(self.get_email("Queued"))

        except Exception as e:
            self.log(f"An error occurred while processing an email: {str(e)} \n {traceback.format_exc()}", tag="red", send_email=True)
            self.move_email(mail, "Errors", "inbox")
            return
        

    def handle_login(self, mail): # Handles login emails, not implemented
        msg = self.get_msg(mail)
        subject = msg["Subject"]
        filepaths = Loginulator.get_filepaths(msg)
        if not filepaths:
            self.log(f"Loginulator failed for '{subject}'", tag="red", send_email=True)
            self.move_email(mail, "Need_Login")
            return True
        
        for filepath in filepaths:
            new_filepath = Rectangulator.main(filepath, self, self.TEMPLATE_FOLDER)
            try:
                # Save invoice
                os.rename(filepath, new_filepath)
                self.log(f"Created new invoice file {os.path.basename(new_filepath)}", tag="blue")
                self.print_invoice(new_filepath, mail)
                return False
            except Exception as e:
                self.log(f"An error occurred while renaming {filepath} to {new_filepath}: {str(e)}", tag="red", send_email=True)
                return True


    def handle_attachments(self, mail): # Iterate over email parts and find pdf
        msg = self.get_msg(mail, "Queued")
        subject = msg["Subject"]
        filenames = []

        for part in msg.walk():
            if part.get_filename() not in filenames and part.get_content_disposition() is not None and part.get_filename() is not None and part.get_filename().lower().endswith(".pdf"):
                filenames.append(part.get_filename())
                if subject == "Test":
                    self.add_to_queue(mail, part, testing=True)
                else:
                    self.add_to_queue(mail, part)
        if subject != "Test":
            self.add_to_queue("End", None)


    def add_to_queue(self, mail, part, testing=False): # Adds invoice to Rectangulator queue
        if mail == "End": # tell rectangulator it's the end of the email (since there could be multiple attachments)
            self.rectangulator_handler.add_to_queue("End", None, None, self, None, None)
            return

        # Get fllename and attachment
        filename = part.get_filename()
        filepath = os.path.join(self.INVOICE_FOLDER, filename)
        attachment = part.get_payload(decode=True)
        
        # Check if file already exists
        if os.path.exists(filepath):
            filename = f"{filename[:-4]}_{str(int(time.time()))}.pdf"
            filepath = os.path.join(self.INVOICE_FOLDER, filename)
        
        # Download invoice PDF
        with open(filepath, "wb") as file:
            file.write(attachment)

        if testing: # when testing inbox
            filename = f"Test_{filename}"
            self.rectangulator_handler.add_to_queue(mail, filename, filepath, self, self.TEMPLATE_FOLDER, self.TESTING)
            return

        # Prompt user to make template, timeout if it takes too long
        self.rectangulator_handler.add_to_queue(mail, filename, filepath, self, self.TEMPLATE_FOLDER, self.TESTING)


    def move_email(self, mail, label, og_label): # Moves email to label
        subject = "Unknown"
        try:
            # Get msg and subject if possible
            msg = self.get_msg(mail, og_label)
            if msg:
                subject = msg["Subject"]

            # Make a copy of the email in the specified label
            copy = self.imap.copy(mail, label)

            # Mark the original email as deleted
            self.imap.store(mail, "+FLAGS", "\\Deleted")
            self.imap.expunge()
            self.log(f"Moved email '{subject}' from {og_label} to {label}.", tag="blue")
            return copy
        except Exception as e:
            self.log(f"Transfer failed for '{subject}': {str(e)}", tag="red", send_email=True)


    def send_email(self, body): # Sends email to me
        sender_email = f"{self.username}{self.ADDRESS}"
        try:
            if self.TESTING:
                return 
                
            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["Subject"] = "Alert"
            message["From"] = sender_email
            message["To"] = self.RECIEVER_EMAIL
            message.attach(MIMEText(body, "plain"))

            # Send the email using SMTP
            with smtplib.SMTP(self.SMTP_SERVER, 587) as server:
                server.starttls()
                server.login(sender_email, self.password)
                server.sendmail(sender_email, self.RECIEVER_EMAIL, message.as_string())
        except Exception as e:
                self.log(f"Error sending email - {str(e)}", tag="red")


    def get_email(self, label): # Gets most recent email in label
        try:
            self.imap.select(label)
            _, data = self.imap.search(None, "ALL")
            email_id = data[0].split()[-1]
            return email_id
        except Exception as e:
            self.log(f"An error occurred while getting email: {str(e)}", tag="red", send_email=True)
            return None


    def log(self, *args, tag=None, send_email=False): # Logs to text box and log file
        try:
            if self.window_closed: # check if window is still open
                return
            message = " ".join([str(arg) for arg in args]) # convert args to string

            # Get rid of no_new_emails messages
            if tag == "no_new_emails":
                self.remove_messages(message)
            
            # Insert the new message to the text widget
            self.log_text_widget.insert(tk.END, message + "\n", (tag, "default"))
            # If the bottom quarter of the text widget is visible, autoscroll
            if self.log_text_widget.yview()[1] > 0.75:
                self.log_text_widget.yview_moveto(1)

            # Send email for errors
            if tag == "red" and send_email:
                self.send_email(message)
            
            # Write to the log file
            with open(self.LOG_FILE, "a") as file:
                file.write(message + "\n")
        except Exception as e:
            print(f"Error logging: {str(e)}")


    def check_labels(self, labels): # Checks for emails that need to be looked at in labels
        for label in labels:
            try:
                # Check if any emails in specified label
                self.imap.select(label)
                _, data = self.imap.search(None, "ALL")
                email_ids = data[0].split()

                # Alert user if there are emails
                if len(email_ids) > 0:
                    self.log(f"{len(email_ids)} emails in {label} - {self.current_time} {self.current_date}", tag="orange")
                if len(labels) == 1:
                    return email_ids
            except Exception as e:
                self.log(f"An error occurred while checking the label: {str(e)}", tag="red", send_email=True)
        
    
    def get_msg(self, mail, label): # Gets email message
        try:
            self.imap.select(label)
            _, data = self.imap.fetch(mail, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            return msg
        except Exception as e:
            self.log(f"Error getting message: {str(e)}", tag="red", send_email=True)
            return None


    def remove_messages(self, message): # Removes no_new_emails messages
        message = message[:-22] # cuts out the date-time

        # Searches for every no_new_emails message then deletes it
        index = self.log_text_widget.search(message, "1.0", tk.END)
        while index:
            self.log_text_widget.delete(index, f"{index}+{len(message) + 23}c") # +1 for new line, +22 for date-time
            index = self.log_text_widget.search(message, "1.0", tk.END)
            self.root.update()


    def resolve_errors(self): # Moves error emails back to inbox
        try:
            self.log(f"Attempting to resolve errors.", tag="blue")
            # Get emails in error label
            email_ids = self.check_labels(["Errors"])

            if len(email_ids) == 0:
                self.log(f"No errors to resolve.", tag="blue")
                return

            # Move emails back to inbox
            for email_id in email_ids:
                self.move_email(email_id, "inbox", "Errors")
        except Exception as e:
            self.log(f"Error resolving errors: {str(e)}", tag="red", send_email=True)


    def resolve_prints(self): # Moves unprinted invoices back to inbox
        try:
            self.log(f"Attempting to resolve unprinted invoices.", tag="blue")
            # Get emails in Need_Print label
            email_ids = self.check_labels(["Need_Print"])

            if len(email_ids) == 0:
                self.log(f"No unprinted invoices to resolve.", tag="blue")
                return

            # Move emails back to inbox
            for email_id in email_ids:
                self.move_email(email_id, "inbox", "Need_Print")
        except Exception as e:
            self.log(f"Error resolving unprinted invoices: {str(e)}", tag="red", send_email=True)


    def pause_processing(self): # Pauses processing
        self.log("Processing paused.", tag="yellow")
        self.pause_button.config(text="Resume", command=self.resume_processing)
        self.errors_button.config(state=tk.DISABLED)
        self.print_errors_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)
        self.pause_event.set()


    def resume_processing(self): # Resumes processing
        self.log("Processing resumed.", tag="yellow")
        self.pause_button.config(text="Pause", command=self.pause_processing)
        self.errors_button.config(state=tk.NORMAL)
        self.print_errors_button.config(state=tk.NORMAL)
        self.test_inbox_button.config(state=tk.NORMAL)
        self.pause_event.clear()

 
    def restart_processing(self): # Restarts processing
        self.log(f"Restarting...", tag="orange")
        self.disconnect()
        self.processor_thread = None
        self.main()

   
    def logout(self): # Logs out
        self.log("Logging out...", tag="orange")
        self.pause_button.config(state=tk.DISABLED)
        self.errors_button.config(state=tk.DISABLED)
        self.logout_button.config(state=tk.DISABLED)
        self.test_inbox_button.config(state=tk.DISABLED)
        self.pause_event.set()
        self.processor_running = False
        self.logging_out = True

   
    def toggle_testing(self): # Toggles testing mode
        if self.TESTING:
            self.TESTING = False
            self.testing_button.config(bg="#FFCCCC")
        else:
            self.TESTING = True
            self.testing_button.config(bg="#CCFFCC")


    def test_rectangulator(self): # Opens rectangulator with test invoice
        self.log("Testing Rectangulator...", tag="orange")
        return_list = []
        self.rectangulator_handler.add_to_queue(None, None, self.TEST_INVOICE, self, self.TEST_TEMPLATE_FOLDER, True)
        if return_list != []:
            new_filepath, should_print = return_list
            self.log(f"new_filepath: {new_filepath}, should_print: {should_print}", tag="orange")
        self.log("Testing complete.", tag="orange")


    def test_inbox(self): # Sends test email to inbox, won't be printed or downloaded
        self.log("Sending test email to inbox", tag="orange")
        mail = self.get_email("Test_Email")
        self.move_email(mail, "inbox", "Test_Email")


    def reconnect(self): # Reconnects to email
        self.disconnect(log=False)
        self.imap = self.connect(log=False)
        self.log(f"Reconnected to {self.username} - {self.current_time} {self.current_date}", tag="green")


    def on_program_exit(self): # Runs when program is closed, disconnects and closes window
        print("Program closed")
        self.log("Disconnecting...", tag="red")
        self.root.update()
        self.window_closed = True

        # Disconnect imaps if running
        if self.processor_thread:
            self.processor_running = False  #set the flag to stop the email processing loop
            self.pause_event.set()
            self.processor_thread.join()

        # Destroys tkinter window 
        self.root.destroy()


    def flash_taskbar(self): # Flash icon in taskbar
        # Code from stack overflow
        hwnd_int = int(self.root.frame(), base=16)
        win32gui.FlashWindow(hwnd_int, 0)


    @property
    def current_time(self):
        return time.strftime("%H:%M:%S", time.localtime())
    

    @property
    def current_date(self):
        return time.strftime("%m-%d-%Y", time.localtime())