import Rectangulator
import win32print
import threading
import traceback
import win32api
import imaplib
import smtplib
import config
import email
import time
import os
import tkinter as tk
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class EmailProcessor:
    # CONSTANTS
    LOG_FILE = config.LOG_FILE
    LOG_FILE2 = config.LOG_FILE2
    TEMPLATE_FOLDER = config.TEMPLATE_FOLDER
    INVOICE_FOLDER = config.INVOICE_FOLDER
    ACP_USER, ACP_PASS = config.ACP_USER, config.ACP_PASS
    APC_USER, APC_PASS = config.APC_USER, config.APC_PASS
    IMAP_SERVER = config.IMAP_SERVER
    SMTP_SERVER = config.SMTP_SERVER
    RECIEVER_EMAIL = config.RECIEVER_EMAIL
    TRUSTED_ADDRESS = config.TRUSTED_ADDRESS
    ADDRESS = config.ADDRESS
    WAIT_TIME = 10 #seconds
    RECONNECT_CYCLE_COUNT = 3600 / WAIT_TIME #1 hour
    TESTING = True

    def __init__(self, root):
        # VARIABLES
        self.alert_window = None #used for pop ups
        self.window_closed = None
        self.processor_thread = None
        self.processor_running = False
        self.pause_event = threading.Event() #used for cycles
        self.root = root
        self.connected = False

        # GUI BUTTONS
        self.button_frame = tk.Frame(root)
        self.button_frame.pack(side=tk.TOP)

        self.start_button = tk.Button(self.button_frame, text="Start Process", command=self.main) #start process button
        self.start_button.pack(side=tk.LEFT, padx=1)

        self.pause_button = tk.Button(self.button_frame, text="Pause", command=self.pause_processing, state=tk.DISABLED) #pause button
        self.pause_button.pack(side=tk.LEFT, padx=1)

        self.resume_button = tk.Button(self.button_frame, text="Resume", command=self.resume_processing, state=tk.DISABLED) #resume button
        self.resume_button.pack(side=tk.LEFT, padx=1)

        self.restart_button = tk.Button(self.button_frame, text="Restart", command=self.restart_processing, state=tk.DISABLED) #restart button
        self.restart_button.pack(side=tk.LEFT, padx=1)

        self.log_text_widget = tk.Text(root, height=30, width=140, spacing1=4, padx=0, pady=0) #text label
        self.log_text_widget.pack()

        # GUI STYLES
        self.log_text_widget.tag_configure("red", background="#FFCCCC")
        self.log_text_widget.tag_configure("yellow", background="yellow")
        self.log_text_widget.tag_configure("orange", background="#FFB434")	
        self.log_text_widget.tag_configure("lgreen", background="#CCFFCC")	
        self.log_text_widget.tag_configure("green", background="#39FF12")	
        self.log_text_widget.tag_configure("dgreen", background="#00994d")	
        self.log_text_widget.tag_configure("blue", background="#89CFF0")
        self.log_text_widget.tag_configure("gray", background="#DEDDDD")
        self.log_text_widget.tag_configure("no_new_emails", background="#DEDDDD") #gray
        self.log_text_widget.tag_configure("default", borderwidth=0.5, relief="solid", lmargin1=10, offset=8) #default

        self.root.protocol("WM_DELETE_WINDOW", self.on_program_exit) #runs exit protocol on window close
        
    def main(self): # Runs when start button is pressed
        self.log("Connecting...", tag="dgreen")
        self.root.update()
        self.processor_running = True

        # Enable and disable buttons
        self.start_button.config(state=tk.DISABLED) 
        self.pause_button.config(state=tk.NORMAL)
        self.restart_button.config(state=tk.NORMAL)
        
        # ACP login
        self.imap_acp = self.connect(self.ACP_USER, self.ACP_PASS)
        if self.imap_acp:
            self.processor_thread = threading.Thread(target=self.search_inbox, args=[self.imap_acp])
            self.processor_thread.start()

        # APC login
        self.imap_apc = self.connect(self.APC_USER, self.APC_PASS)
        if self.imap_apc:
            self.processor_thread = threading.Thread(target=self.search_inbox, args=[self.imap_apc])
            self.processor_thread.start()

    def connect(self, username, password, log=True): # returns imap object
        user = f"{username}{self.ADDRESS}"
        try:
            # Log into email
            imap = imaplib.IMAP4_SSL(self.IMAP_SERVER)
            imap.login(user, password)
            self.connected = True
            if log:
                self.log(f"--- Connected to {username} --- {self.current_time} {self.current_date}", tag="dgreen")
            return MYImap(imap, username, password)
        except imaplib.IMAP4_SSL.error as e:
            self.log(f"Unable to connect to {username}: {str(e)}", tag="red", sender_imap=imap)
            return
        
    def disconnect(self, imap, log=True):
        try:
            # Logout
            imap.imap.logout()
            self.connected = False
            if log:
                self.log(f"--- Disconnected from {imap.username} --- {self.current_time} {self.current_date}", tag="red")
        except Exception as e:
            self.log(f"An error occurred while disconnecting: {str(e)}", tag="red", sender_imap=imap)        

    def search_inbox(self, imap): # This is what runs each cycle
        try:
            cycle_count = 0
            while self.processor_running:
                if not self.pause_event.is_set() and self.connected:
                    # Search for all emails in the inbox
                    imap.imap.select("inbox")
                    _, emails = imap.imap.search(None, "ALL")

                    # Check if no new mail
                    if not emails[0].split():
                        self.log(f"No new emails for {imap.username} - {self.current_time} {self.current_date}", tag="no_new_emails")

                        # Check if emails need to be looked at
                        self.check_labels(imap, ["Need_Print", "Need_Login", "Errors"])

                        # Pause until next cycle
                        self.pause_event.wait(timeout=self.WAIT_TIME)
                    else:
                        self.process_email(imap, emails[0].split()[0])

                    cycle_count += 1
                    # Reconnect every hour
                    if cycle_count == self.RECONNECT_CYCLE_COUNT:
                        imap = self.reconnect(imap)
                        cycle_count = 0
            # Disconnect when the program is closed
            self.disconnect(imap)
        except Exception as e:  
            self.log(f"An error occurred while searching the inbox for {imap.username}: {str(e)}", tag="red", sender_imap=imap)

    def process_email(self, imap, mail): # Handles each email
        subject = ""
        try: 
            # Fetch email
            _, data = imap.imap.fetch(mail, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            subject = msg["Subject"]
            sender_email = email.utils.parseaddr(msg["From"])[1]

            # Check if sender is trusted
            if not sender_email.endswith(self.TRUSTED_ADDRESS):
                self.move_email(imap, mail, "Not_Invoices", subject)
                return
            
            # Check for attachments
            has_attachment = any(part.get("content-disposition", "").startswith("attachment") for part in msg.walk() if msg.is_multipart())
            if not has_attachment:
                self.log(f"'{subject}' has no attachment, assuming login needed for {imap.username}", tag="yellow")
                self.move_email(imap, mail, "Need_Login", subject)
                return 
            
            # Handle attachments
            attachment_error = self.handle_attachments(imap, mail, msg, subject)

            # Move to invoices label if no errors
            if not attachment_error:
                self.move_email(imap, mail, "Invoices", subject)
            else:
                self.log(f"'{subject}' failed to download, moved to Error label for {imap.username}", tag="red", sender_imap=imap)
                self.move_email(imap, mail, "Errors", subject)

        except Exception as e:
            self.log(f"An error occurred while processing an email for {imap.username}: {str(e)} \n {traceback.format_exc()}", tag="red", sender_imap=imap)
            self.move_email(imap, mail, "Errors", subject)
            return
        
    def handle_attachments(self, imap, mail, msg, subject):        # Iterate over email parts and find pdf
        error = False
        for part in msg.walk():
            if part.get_content_disposition() is not None and part.get_filename() is not None and part.get_filename().lower().endswith(".pdf"):
                # Check if download is successful
                invoice_downloaded, filepath = self.download_invoice(part, imap)
                if invoice_downloaded == "not_invoice":
                    continue
                elif not invoice_downloaded:
                    error = True
                    continue
            
                if not self.TESTING:
                    # Check if print is successful
                    invoice_printed = self.print_invoice(filepath, imap)
                    if not invoice_printed:
                        self.move_email(imap, mail, "Need_Print", subject)
                        continue
        return error

    def download_invoice(self, part, imap):
        # Get fllename and attachment
        filename = part.get_filename()
        attachment = part.get_payload(decode=True)
        
        filepath = os.path.join(self.INVOICE_FOLDER, filename)
        
        # Check if file already exists
        if os.path.exists(filepath):
            self.log(f"Invoice file already exists at {filepath} for {imap.username}", tag="red", sender_imap=imap)
            return False, None

        # Download invoice PDF
        with open(filepath, 'wb') as file:
            file.write(attachment)

        # Prompt user to draw rectangles
        new_filepath = Rectangulator.main(filepath, self.TEMPLATE_FOLDER, self.LOG_FILE2, self)

        # Check if not invoice
        if new_filepath == "not_invoice":
            os.remove(filepath)
            return "not_invoice", None

        # Check if Rectangulator fails
        if new_filepath == None:
            self.log(f"Rectangulator failed for {imap.username}", tag="red", sender_imap=imap)
            os.remove(filepath)
            return False, None
        
        # Check if invoice has already been processed
        if os.path.exists(new_filepath):
            os.remove(filepath)
            self.log(f"New invoice file already exists at {new_filepath} for {imap.username}", tag="red", sender_imap=imap)
            return False, None
        
        # Save invoice
        os.rename(filepath, new_filepath)
        self.log(f"Created new invoice file {os.path.basename(new_filepath)} for {imap.username}", tag="blue")
        return True, new_filepath
        
    def print_invoice(self, filepath, imap): # Printer
        try:
            # Get default printer and print
            p = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", filepath, None,  ".",  0)
            self.log(f"Printed {filepath} completed successfully for {imap.username}.", tag="blue")
            return True
        except Exception as e:
            self.log(f"Printing failed: {str(e)}", tag="red", sender_email=imap.username)
            return False

    def move_email(self, imap, mail, label, subject): # Moves emails to labels
        try:
            # Make a copy of the email in the specified label
            copy = imap.imap.copy(mail, label)

            # Mark the original email as deleted
            imap.imap.store(mail, '+FLAGS', '\\Deleted')
            imap.imap.expunge()
            self.log(f"Email '{subject}' moved to {label} for {imap.username}.", tag="blue")
        except Exception as e:
            self.log(f"Email '{subject}' transfer failed for {imap.username}: {str(e)}", tag="red", sender_imap=imap)

    def send_email(self, imap, subject, body):
        try:
            sender_email = f"{imap.username}{self.ADDRESS}"

            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["Subject"] = subject
            message["From"] = sender_email
            message["To"] = self.RECIEVER_EMAIL
            message.attach(MIMEText(body, "plain"))

            # Send the email using SMTP
            with smtplib.SMTP(self.SMTP_SERVER, 587) as server:
                server.starttls()
                server.login(sender_email, imap.password)
                server.sendmail(sender_email, self.RECIEVER_EMAIL, message.as_string())
                self.log(f"Email from {imap.username} sent to {self.RECIEVER_EMAIL}", tag="gray")
        except Exception as e:
                self.log(f"Error sending email from {imap.username} - {str(e)}", tag="red")

    def log(self, *args, tag=None, sender_imap=None): # Logs to text box and log file
        try:
            if self.window_closed: #check if window is still open
                return
            message = ' '.join([str(arg) for arg in args]) #convert args to string

            # Get rid of no_new_emails messages
            if tag == "no_new_emails":
                self.remove_messages(message)
            
            # Insert the new message to the text widget
            self.log_text_widget.insert(tk.END, message + "\n", (tag, "default"))
            self.log_text_widget.see(tk.END)  #scroll to the end of the text widget   

            # Send email for errors
            if tag == "red" and sender_imap:
                self.send_email(sender_imap, "Error Alert", message)
            
            # Write to the log file
            with open(self.LOG_FILE, "a") as file:
                file.write(message + "\n")
        except Exception as e:
            print(f"Error logging: {str(e)}")

    def check_labels(self, imap, labels): # Checks for emails that need to be looked at in labels
        for label in labels:
            try:
                # Check if any emails in specified label
                imap.imap.select(label)
                _, data = imap.imap.search(None, 'ALL')
                email_ids = data[0].split()

                # Alert user if there are emails
                if len(email_ids) > 0:
                    self.log(f"{len(email_ids)} emails in {label} for {imap.username} - {self.current_time} {self.current_date}", tag="orange")
            except Exception as e:
                self.log(f"An error occurred while checking the label for {imap.username}: {str(e)}", tag="red", sender_imap=imap)
    
    def remove_messages(self, message): # Removes no_new_emails messages
        message = message[:-22] #cuts out the date-time

        # Searches for every no_new_emails message then deletes it
        index = self.log_text_widget.search(message, "1.0", tk.END)
        while index:
            self.log_text_widget.delete(index, f"{index}+{len(message)+1+22}c") #+1 for new line, +22 for date-time
            index = self.log_text_widget.search(message, "1.0", tk.END)
            self.root.update()

    def pause_processing(self): # Pauses processing
        self.pause_event.set()
        self.log("Processing paused.", tag="yellow")
        self.pause_button.config(state=tk.DISABLED)
        self.resume_button.config(state=tk.NORMAL)
        self.restart_button.config(state=tk.DISABLED)

    def resume_processing(self): # Resumes processing
        self.pause_event.clear()
        self.log("Processing resumed.", tag="yellow")
        self.pause_button.config(state=tk.NORMAL)
        self.resume_button.config(state=tk.DISABLED)
        self.restart_button.config(state=tk.NORMAL)

    def restart_processing(self): # Restarts processing
        self.log("Restarting...", tag="orange")
        self.processor_running = False
        self.main()

    def reconnect(self, imap): # Reconnects to imap
        self.disconnect(imap, log=False)
        imap = self.connect(imap.username, imap.password, log=False)
        self.log(f"Reconnected to {imap.username} - {self.current_time} {self.current_date}", tag="green")
        return imap

    def on_program_exit(self): # Runs when program is closed
        self.log("Disconnecting...", tag="red")
        self.root.update()
        self.window_closed = True

        # Close alert windows
        if self.alert_window:
            self.alert_window.destroy() 

        # Disconnect imaps if running
        if self.processor_thread:
            self.processor_running = False  #set the flag to stop the email processing loop
            self.pause_event.set()
            self.processor_thread.join()

        # Destroys tkinter window 
        self.root.destroy()

    def show_alert(self, *args): # Create Yes/No popup window
        message = ' '.join([str(arg) for arg in args])
        self.alert_window = AlertWindow(self.root, message)
        self.root.wait_window(self.alert_window)
        return self.alert_window.choice

    @property
    def current_time(self):
        return time.strftime("%H:%M:%S", time.localtime())
    
    @property
    def current_date(self):
        return time.strftime("%Y-%m-%d", time.localtime())

class MYImap:
    
    def __init__(self, imap, username, password):
        self.imap = imap
        self.username = username
        self.password = password


class AlertWindow(tk.Toplevel):

    def __init__(self, parent, message):
        super().__init__(parent)
        self.title("Alert")
        self.geometry("300x100")
        self.choice = None

        label = tk.Label(self, text=message, wraplength=250) # text label
        label.pack(padx=20, pady=20)

        yes_button = tk.Button(self, text="Yes", command=self.on_yes_button_click) # yes button
        yes_button.pack(side=tk.LEFT)

        no_button = tk.Button(self, text="No", command=self.on_no_button_click) # no button
        no_button.pack(side=tk.LEFT)

    def on_yes_button_click(self):
        self.choice = True
        self.destroy()

    def on_no_button_click(self):
        self.choice = False
        self.destroy()


if __name__ == "__main__":
    # Get app icon       
    icon_path = os.path.join(os.path.dirname(__file__), "hotpot.ico")   

    # Setup gui
    root = tk.Tk()
    root.title("Pewter")
    root.iconbitmap(icon_path)
    root.geometry("1200x600")
    email_processor = EmailProcessor(root)
    root.mainloop()