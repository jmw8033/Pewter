from email.mime.multipart import MIMEMultipart
from matplotlib.widgets import TextBox
from matplotlib.patches import Rectangle
from matplotlib.widgets import Button
from matplotlib.widgets import CheckButtons
from email.mime.text import MIMEText
from Alertinator import AlertWindow
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
import win32print
import win32api
import traceback
import threading
import warnings
import smtplib
import config
import email
import time
import fitz
import glob
import re
import os

warnings.simplefilter("ignore", UserWarning)
class RectangulatorHandler:

    def __init__(self):
        self.queue = []
        self.invoice = True
        self.log_file = config.LOG_FILE
        self.should_print = True
        self.hit_submit = False
        self.root = None
        self.queue_loop = threading.Thread(target=self.process_queue)
        self.queue_loop.start()
        self.current_email = None
        self.current_email_dest = None


    def add_to_queue(self, mail, filename, filepath, root, template_folder, return_list, testing=False): # Add a file to the queue, or process it immediately if template exists
        if mail == "End": # signal to end the current email and move to correct label
            self.queue.append([mail, filename, filepath, root, template_folder, return_list, testing])
            return

        if testing and mail == None: # when specifically pressing the test rectangulator button
            self.rectangulate(filename, filepath, root, template_folder, return_list, testing)
            return            

        # check if template exists and use it
        template_exists = self.check_templates(filepath, template_folder, return_list, root)
        if template_exists:
            new_filepath = template_exists[0]
            self.current_email = mail
            self.set_dest_label("Invoices")
            if os.path.exists(new_filepath): # check if file already exists
                self.log(f"New invoice file already exists at {filepath}", tag="orange", root=root)
                new_filepath = f"{filepath[:-4]}_{str(int(time.time()))}.pdf"
            os.rename(filepath, new_filepath)
            self.log(f"Created new invoice file {os.path.basename(new_filepath)} - {root.current_date} {root.current_time}", tag="blue", root=root)
            if not testing:
                self.print_invoice(new_filepath, root)
            return

        # otherwise add to queue
        self.queue.append([mail, filename, filepath, root, template_folder, return_list, testing])
        self.log(f"Template required for {filename}  --- {root.current_time} {root.current_date}", root=root)

    
    def process_queue(self): # Main loop for processing the queue
        while True:
            try:
                if len(self.queue) == 0:
                    time.sleep(config.QUEUE_CYCLE_TIME)
                    continue
                
                mail, filename, filepath, root, template_folder, return_list, testing = self.queue.pop(0)
                if mail == "End":
                    self.move_email(self.current_email, self.current_email_dest, "Queued", root)
                    self.current_email = None
                    self.current_email_dest = None
                    continue
                else:
                    self.current_email = mail

                return_list = self.rectangulate(filename, filepath, root, template_folder, return_list, testing)

                # Check if Rectangulator fails
                if return_list == [] or return_list[0] == None:
                    self.log(f"Rectangulator failed", tag="red", send_email=True, root=root)
                    os.remove(filepath)
                    subject = self.get_msg(mail, "Queued", root)["Subject"]
                    self.log(f"'{subject}' failed to download, moved to Error label", tag="red", send_email=True, root=root)
                    self.set_dest_label("Errors")
                    continue
                
                new_filepath, should_print = return_list
                if testing:
                    should_print = False

                # Check if not invoice
                if new_filepath == "not_invoice":
                    self.log(f"'{filename}' is not an invoice", tag="blue", root=root)
                    if should_print:
                        self.print_invoice(filepath, root)
                    os.remove(filepath)
                    self.set_dest_label("Not_Invoices")
                    continue
                
                # Check if invoice has already been processed
                if os.path.exists(new_filepath):
                    old_filepath = new_filepath
                    self.log(f"New invoice file already exists at {os.path.basename(old_filepath)}", tag="orange", root=root)
                    new_filepath = f"{old_filepath[:-4]}_{str(int(time.time()))}.pdf"
                
                # Save invoice
                os.rename(filepath, new_filepath)
                self.set_dest_label("Invoices")
                self.log(f"Created new invoice file {os.path.basename(new_filepath)} - {root.current_date} {root.current_time}", tag="blue", root=root)
                self.print_invoice(new_filepath, root)

            except Exception as e:
                self.log(f"An error occurred while processing the queue: {str(e)}", tag="red", send_email=True, root=root)

    
    def rectangulate(self, filename, filepath, root, template_folder, return_list, testing=False): # Main function for the Rectangulator
        self.root = root

        # If no template exists, make one
        try:
            if not testing:
                self.send_email("Must create template", root) # email me

            rectangulator, text_box = self.setup_page(filepath, template_folder) # setup all elements on page
        
            if not self.invoice:#  if the user clicked the "Not An Invoice" button
                self.invoice = True
                return_list.extend(["not_invoice", self.should_print])
                return return_list
            filename = rectangulator.rename_pdf()
            if filename: # if the user dragged a rectangle
                return_list.extend([filename, self.should_print])
                return return_list
            elif text_box.text: # if the user entered a filename
                return_list.extend([os.path.join(os.path.dirname(filepath), f"{text_box.text}.pdf"), self.should_print])
                return return_list

        except Exception as e:
            self.log(f"An error occurred while drawing rectangles: {str(e)}", tag="red", root=root)

        return_list = [None, False]
        return return_list


    def check_templates(self, pdf_path, template_folder, return_list, root): # Check if a template exists for the invoice
        for file in glob.glob(rf"{template_folder}\*.txt"):
            try:
                with open(file, "r") as f:
                    # Get the invoice name, date, and number from the template
                    invoice_name = f.readline().split("?")
                    invoice_date = f.readline().split("?")
                    invoice_num = f.readline().split("?")
                    # Get the company name from the invoice
                    identifier = self.sanitize_filename(self.get_text_in_rect(Rectangle((invoice_name[1], invoice_name[2]), invoice_name[3], invoice_name[4]), pdf_path))

                    # If company name on invoice matches name on template, use that template
                    if invoice_name[0] == identifier:
                        self.log(f"Used template {file} for {identifier}", root=root)
                        # Get the invoice date and number from the invoice
                        invoice_date = self.get_text_in_rect(Rectangle((invoice_date[1], invoice_date[2]), invoice_date[3], invoice_date[4]), pdf_path)
                        invoice_num = self.get_text_in_rect(Rectangle((invoice_num[1], invoice_num[2]), invoice_num[3], invoice_num[4]), pdf_path)
                        # Clean the invoice date
                        invoice_date = self.check_date_outlier(invoice_name[0], invoice_date).replace("/", "-")
                        return_list.extend([rf"{os.path.dirname(pdf_path)}\{invoice_date}_{invoice_num}.pdf"])
                        return return_list
            except Exception as e:
                pass


    def setup_page(self, pdf_path, template_folder): # Setup the page for the Rectangulator
            doc = fitz.open(pdf_path)
            page = doc[0]

            # Convert the page to a NumPy array for plotting
            pix = page.get_pixmap()
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

            # Create a matplotlib figure and axis for displaying the image
            fig, ax = plt.subplots(figsize=(9, 9))
            plt.subplots_adjust(bottom=0.2)
            ax.imshow(img_array)

            # Create a Not An Invoice button
            button = Button(plt.axes([0.65, 0.05, 0.2, 0.075]), "Not An Invoice")
            button.on_clicked(self.not_invoice)

            # Create a checkbox for if it should be printed
            printCheckBox = CheckButtons(plt.axes([0.9, 0.065, 0.03, 0.03]), [""], [True])
            def print_callback(label):
                self.should_print = not self.should_print
            printCheckBox.on_clicked(print_callback)
            for i, line in enumerate(printCheckBox.lines):
                rect = printCheckBox.rectangles[i]
                rect.set_width(0.5)
                rect.set_height(0.5)
                rect.set_edgecolor("none")
                # Calculate the center of the rectangle
                center_x = rect.get_x() + rect.get_width() / 2
                center_y = rect.get_y() + rect.get_height() / 2
                # Update the line positions to be centered
                line[0].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
                line[1].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
                line[0].set_ydata([center_y - rect.get_height() / 4, center_y + rect.get_height() / 4])
                line[1].set_ydata([center_y + rect.get_height() / 4, center_y - rect.get_height() / 4])
            printLabel = fig.text(0.896, 0.1, "Print?", fontsize=10)

            # Create a text box to manually enter filename
            text_box = TextBox(plt.axes([0.1, 0.05, 0.45, 0.075]), label="", initial="")
            def on_text_submit(text):
                if self.hit_submit:
                    return
                self.hit_submit = True
                filename_is_correct = AlertWindow(f"Is '{text_box.text}' the correct filename?").get_answer()
                self.hit_submit = False
                if filename_is_correct:
                    plt.close()
            text_box.on_submit(on_text_submit)

            # Create a submit button for the text box
            submit_button = Button(plt.axes([0.45, 0.05, 0.15, 0.075]), "Submit")
            submit_button.on_clicked(on_text_submit)

            # Create text labels for instructions and text box
            text_box_label = fig.text(0.2, 0.14, "Enter Filename (mm-dd-yy_invoice#)", fontsize=10)
            instruction_label = fig.text(0.25, 0.94, "- Draw boxes around Company Name, Date, and Invoice (in that order)", fontsize=10)
            instruction_label_2 = fig.text(0.25, 0.92, "- Company Name can be any piece of text unique to that vendor", fontsize=10)
            instruction_label_3 = fig.text(0.25, 0.90, "- Right click to verify and save", fontsize=10)

            # Create an instance of the Rectangulator and bind it to the axis
            rectangulator = Rectangulator(ax, fig, pdf_path, template_folder, self)

            # Create a timer to close the plot after a set time
            timer = fig.canvas.new_timer(interval=config.RECTANGULATOR_TIMEOUT)
            timer.add_callback(lambda: plt.close())
            timer.start()

            # Show the plot 
            plt.show()
            return rectangulator, text_box


    def print_invoice(self, filepath, root): # Printer
        try:
            # Get default printer and print
            p = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", filepath, None,  ".",  0)
            self.log(f"Printed {os.path.basename(filepath)} completed successfully.", tag="blue", root=root)
            return True
        except Exception as e:
            self.set_dest_label("Need_Print")
            self.log(f"Printing failed for {filepath}: {str(e)}", tag="red", send_email=True, root=root)
            return False


    def set_dest_label(self, label): # Set the destination label for the current email

        if self.current_email_dest not in {"Errors", "Need_Print", "Not_Invoice"}:
            self.current_email_dest = label


    def log(self, *args, tag="purple", send_email=False, root=None): # Log messages to a file and optionally send an email
        with open(self.log_file, "a") as file:
            message = "--- RECTANGULATOR --- " + " ".join([str(arg) for arg in args]) 
            if root:
                root.log(message, tag=tag, send_email=send_email)
            elif self.root:
                self.root.log(message, tag=tag, send_email=send_email)
            if send_email:
                self.send_email(message, root)
            file.write("\n".join([str(arg) for arg in args]))
            print(message)


    def move_email(self, mail, label, og_label, root): # Moves email to label
        subject = "Unknown"
        try:
            # Get msg and subject if possible
            root.imap.select(og_label)
            msg = self.get_msg(mail, og_label, root)
            if msg:
                subject = msg["Subject"]

            # Make a copy of the email in the specified label
            copy = root.imap.copy(mail, label)

            # Mark the original email as deleted
            root.imap.store(mail, "+FLAGS", "\\Deleted")
            root.imap.expunge()
            self.log(f"Email '{subject}' moved to {label}.", tag="blue", root=root)
        except Exception as e:
            self.log(f"Email '{subject}' transfer failed: {str(e)}", tag="red", send_email=True, root=root)


    def get_msg(self, mail, label, root): # Get the message from the specified email
        try:
            root.imap.select(label)
            _, data = root.imap.fetch(mail, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            return msg
        except Exception as e:
            self.log(f"Error getting message: {str(e)}", root=root, tag="red", send_email=True)
            return None


    def get_email(self, label, root): # Get the most recent email from the specified label
        try:
            root.imap.select(label)
            _, data = root.imap.search(None, "ALL")
            mail = data[0].split()[-1]
            return mail
        except Exception as e:
            self.log(f"Error getting email from {label}: {str(e)}", tag="red", root=root)
            return


    def sanitize_filename(self, filename): # Remove invalid characters from the filename
        sanitized_filename = re.sub(r"[^\w_. -]", "", filename.replace("/", "-"))
        return sanitized_filename.strip()


    def get_text_in_rect(self, rect, pdf_path): # Get the text within a specified rectangle in a PDF document
        try:
            # Retrieve the text within the specified rectangle
            x = float(rect.get_x())
            y = float(rect.get_y())
            width = float(rect.get_width())
            height = float(rect.get_height())

            # Open the first page of the PDF document and get words
            doc = fitz.open(pdf_path)
            page = doc[0]
            words = page.get_text("words")

            # Find words that fall within the specified rectangle
            extracted_text = ""
            for word in words:
                word_x, word_y, _, _ = word[:4] # get word coordinates
                if x <= float(word_x) <= x + width and y <= float(word_y) <= y + height:
                    extracted_text += word[4] + " " # append the word to extracted text
            extracted_text = extracted_text.strip() # remove leading/trailing spaces

            if extracted_text == None:
                extracted_text = ""

            return extracted_text
        except Exception as e:
            self.log(f"An error occurred while processing the PDF: {str(e)} {traceback.format_exc()}")
            return ""
        finally:
            doc.close()


    def check_date_outlier(self, invoice_name, invoice_date): # Check if the date is an outlier and correct it
        calendar = {"Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04", "May": "05", "Jun": "06", 
                    "Jul": "07", "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12"}
        calendar2 = {"January": "01", "February": "02", "March": "03", "April": "04", "May": "05", "June": "06", 
                    "July": "07", "Aug": "08", "September": "09", "October": "10", "November": "11", "December": "12"}
        
        if invoice_name == "BUZZI UNICEM USA - Cement":
            # Uses format "DD-Month-YY"
            invoice_date = invoice_date.split("-")
            day = invoice_date[0]
            invoice_date[0] = calendar[invoice_date[1]]
            invoice_date[1] = day
            return "-".join(invoice_date)
        elif invoice_name in {"Alight Solutions LLC", "Compliance Management International", "Taylor Northeast Inc."}:
            # Uses format "Month DD, YYYY"
            invoice_date = invoice_date.replace(",", "").split(" ")
            invoice_date[0] = calendar2[invoice_date[0]]
            invoice_date = "/".join(invoice_date)
        elif invoice_name in {"ADP SCREENING  SELECTION SERVICES", "Muka Development Group Llc"}:
            # Uses format "Mon DD, YYYY"
            invoice_date = invoice_date.replace(",", "").split(" ")
            invoice_date[0] = calendar[invoice_date[0]]
            invoice_date = "/".join(invoice_date)
        
        return self.clean_date(invoice_date.strip())


    def clean_date(self, invoice_date): # Clean the date to be in the format "MM/DD/YY"
        new_date = invoice_date
        try:
            date = datetime.strptime(new_date, "%m/%d/%y")
            new_date = date.strftime("%m/%d/%y")
        except ValueError:
            try:
                date = datetime.strptime(new_date, "%m/%d/%Y")
                new_date = date.strftime("%m/%d/%y")
            except ValueError:
                self.log(f"Could not convert {invoice_date} to date")
                return invoice_date
        if new_date != invoice_date:
            self.log(f"Changed date from {invoice_date} to {new_date}")
        return new_date
        

    def send_email(self, body, root): # Email me
        if root.TESTING:
            return
        
        try:
            sender_email = f"{self.root.username}{config.ADDRESS}"

            # Create a multipart message
            message = MIMEMultipart()
            message["Subject"] = "Alert"
            message["From"] = sender_email
            message["To"] = config.RECIEVER_EMAIL
            message.attach(MIMEText(body, "plain"))

            # Send the email
            with smtplib.SMTP(config.SMTP_SERVER, 587) as server:
                server.starttls()
                server.login(sender_email, root.password)
                server.sendmail(sender_email, config.RECIEVER_EMAIL, message.as_string())
                self.log(f"Template request sent from {sender_email} to {config.RECIEVER_EMAIL}")
        except Exception as e:
                self.log(f"Error sending email from {sender_email} - {str(e)}")


    def not_invoice(self, event): # If the user clicks the "Not Invoice" button
        self.invoice = False
        plt.close()


class Rectangulator:

    def __init__(self, ax, fig, pdf_path, template_folder, rectangulator_handler):
        # VARIABLES
        self.rectangulator_handler = rectangulator_handler
        self.log_file = rectangulator_handler.log_file
        self.pdf_path = pdf_path
        self.template_folder = template_folder
        self.fig = fig
        self.ax = ax

        self.rectangles = [] # contains rectangle objects
        self.coordinates = [] # contains coordinates of rectangle objects
        self.correcting_rect_index = None
        self.start_x = None
        self.start_y = None
        self.rect = None 
        self.zoom_factor = 1.2
        self.pan_factor = 1
        self.pan_start = None
        self.prev_x = None
        self.prev_y = None
        self.initial_xlim = self.ax.get_xlim()
        self.initial_ylim = self.ax.get_ylim()

        # Connect the event handlers to the canvas
        self.ax.figure.canvas.mpl_connect("button_press_event", self.on_button_press)
        self.ax.figure.canvas.mpl_connect("button_release_event", self.on_button_release)
        self.ax.figure.canvas.mpl_connect("motion_notify_event", self.on_move)
        self.ax.figure.canvas.mpl_connect("scroll_event", self.on_scroll)
        self.ax.figure.canvas.mpl_connect("key_press_event", self.on_key_press)  # Add key press event handler


    def rename_pdf(self):
        try:
            if len(self.rectangles) == 3:
                self.save_template()
            
                # Rename the PDF based on extracted text from rectangles in format "MM-DD-YY_INVOICE_NUMBER"
                extracted_texts = [self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path) for rect in self.rectangles if self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path)]
                extracted_texts[1] = self.rectangulator_handler.check_date_outlier(extracted_texts[0], extracted_texts[1])
                extracted_text_combined = "_".join(extracted_texts[1:])
                sanitized_extracted_text = self.rectangulator_handler.sanitize_filename(extracted_text_combined)
                new_filename = os.path.join(os.path.dirname(self.pdf_path), f"{sanitized_extracted_text}.pdf")
                return new_filename 
        except RecursionError:
            self.rectangulator_handler.log("Window closed please try again")
            return None
        except Exception as e:
            self.rectangulator_handler.log(traceback.format_exc())
            self.rectangulator_handler.log(f"Error occurred with download, {str(e)}")
            return None
        

    def save_template(self):
        # Save the template to a text file as Company Name?x?y?width?height and so on
        for i, coord in enumerate(self.coordinates):
            x, y, width, height = coord
            rect_text = self.rectangulator_handler.get_text_in_rect(self.rectangles[i], self.pdf_path)
            text = f"{rect_text}?{x}?{y}?{width}?{height}\n"
            filename = rf"{self.template_folder}\{self.rectangulator_handler.sanitize_filename(self.rectangulator_handler.get_text_in_rect(self.rectangles[0], self.pdf_path))}.txt"
            with open(filename, "a") as file:
                if i == 0:
                    file.write(f"{self.rectangulator_handler.sanitize_filename(rect_text)}?{x}?{y}?{width}?{height}\n")
                else:
                    file.write(text)
        self.rectangulator_handler.log(f"Created vendor invoice template {self.rectangulator_handler.get_text_in_rect(self.rectangles[0], self.pdf_path)}")


    def on_key_press(self, event):
        if event.key == "escape": # Reset zoom and position
            self.ax.set_xlim(self.initial_xlim)
            self.ax.set_ylim(self.initial_ylim)
            self.pan_start = None
            self.ax.figure.canvas.draw()


    def on_button_press(self, event): # Handle left and right mouse button press events
        if event.button == 1:  # Left mouse button
            # Ignore if the mouse click is outside the plot area
            if event.xdata is None or event.ydata is None:
                return 
            
            self.start_x = event.xdata
            self.start_y = event.ydata
            self.rect = Rectangle((self.start_x, self.start_y), 0, 0, edgecolor="red", linewidth=2, fill=False)
            self.ax.add_patch(self.rect)
            self.ax.figure.canvas.draw()

        elif event.button == 2: # Middle mouse button
            self.pan_start = (event.x, event.y)

        elif event.button == 3:  # Right mouse button
            # Check if 3 rectangles have been drawn
            if len(self.rectangles) != 3:
                self.rectangulator_handler.log("Please draw exactly three rectangles")
                self.reset_rectangles()
                return
            
            # Ask for verifictaion
            def handle_user_input():
                if self.correcting_rect_index is not None:
                    corrected_rect = self.rectangles.pop()
                    corrected_coord = self.coordinates.pop()
                    self.rectangles.insert(self.correcting_rect_index, corrected_rect)
                    self.coordinates.insert(self.correcting_rect_index, corrected_coord)

                 # Show extracted text for verification
                headers = ["--- Company Name: ", "--- Invoice Date: ", "--- Invoice Number: "]
                extracted_text = ""
                for i, rect in enumerate(self.rectangles):
                    extracted_text += (headers[i] + self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path) + "\n")
                text_is_correct = AlertWindow(f"Does the following text match what you selected?\n\n{extracted_text}", 3).get_answer()

                # If user says yes, close the window
                if isinstance(text_is_correct, int) and not isinstance(text_is_correct, bool):
                    self.correcting_rect_index = text_is_correct
                    self.rectangulator_handler.log(f"Please reselect {headers[self.correcting_rect_index]}")
                    self.reset_rectangles(specific_rect=self.correcting_rect_index)
                elif text_is_correct:
                    plt.close(self.fig)
                else:
                    self.rectangulator_handler.log("Please reselect rectangles")
                    self.reset_rectangles()

            input_thread = threading.Thread(target=handle_user_input)
            input_thread.start()
         

    def on_button_release(self, event): # Save rectangle coordinates
        if event.button == 1 and self.rect:  # Left mouse button
            self.start_x = None
            self.start_y = None

            # Append rectangle and coordinated to list
            self.rectangles.append(self.rect)
            self.coordinates.append((self.rect.get_x(), self.rect.get_y(), 
                                    self.rect.get_width(), self.rect.get_height()))  # Store coordinates
            self.rect = None
            self.ax.figure.canvas.draw()

        elif event.button == 2: # Middle mouse button
            self.pan_start = None
            self.prev_x = None
            self.prev_y = None


    def on_move(self, event): # Continuously update drawn rectangle
        # Pan if middle mouse button is pressed
        if event.button == 2 and self.pan_start:
            if self.prev_x is not None and self.prev_y is not None:
                dx = (event.x - self.prev_x) * self.pan_factor
                dy = (event.y - self.prev_y) * self.pan_factor

                xlim = self.ax.get_xlim()
                ylim = self.ax.get_ylim()
                new_xlim = xlim[0] - dx, xlim[1] - dx
                new_ylim = ylim[0] + dy, ylim[1] + dy

                self.ax.set_xlim(new_xlim)
                self.ax.set_ylim(new_ylim)
                self.ax.figure.canvas.draw_idle()
            
            self.prev_x = event.x
            self.prev_y = event.y
            return
        
        if self.rect is None or self.start_x is None or self.start_y is None:
            return

        current_x = event.xdata
        current_y = event.ydata
        # If outside plot area, set to 0
        if current_x is None:
            current_x = 0
        if current_y is None:
            current_y = 0

        # Get rectangle width and height
        width = current_x - self.start_x
        height = current_y - self.start_y

        # Update rectangle
        self.rect.set_width(width)
        self.rect.set_height(height)
        self.ax.figure.canvas.draw()


    def on_scroll(self, event): # Zoom in and out
        if event.button == "down":
            self.zoom(event.xdata, event.ydata, 1 / self.zoom_factor)
        elif event.button == "up":
            self.zoom(event.xdata, event.ydata, self.zoom_factor)


    def reset_rectangles(self, specific_rect=None): # Reset the current rectangles
        if specific_rect is not None:
            self.rectangles[specific_rect].remove()
            self.rectangles.pop(specific_rect)
            self.coordinates.pop(specific_rect)
            self.ax.figure.canvas.draw()
        else:
            for rect in self.rectangles:
                rect.remove()
            self.rectangles = []
            self.coordinates = []
            self.rect = None  
            self.correcting_rect_index = None
            self.ax.figure.canvas.draw()


    def zoom(self, x, y, zoom_factor):
        # Perform zooming based on scroll event
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        if xlim is None or ylim is None:
            return

        if x is None:
            x = np.mean(xlim)
        if y is None:
            y = np.mean(ylim)

        new_xlim = ((xlim[0] - x) / zoom_factor) + x, ((xlim[1] - x) / zoom_factor) + x
        new_ylim = ((ylim[0] - y) / zoom_factor) + y, ((ylim[1] - y) / zoom_factor) + y

        self.ax.set_xlim(new_xlim)
        self.ax.set_ylim(new_ylim)
        self.ax.figure.canvas.draw()


if __name__ == "__main__":
    pass