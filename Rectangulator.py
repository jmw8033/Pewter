from email.mime.multipart import MIMEMultipart
from matplotlib.widgets import TextBox
from matplotlib.patches import Rectangle
from matplotlib.widgets import Button
from matplotlib.widgets import CheckButtons
from email.mime.text import MIMEText
from Alertinator import AlertWindow
from datetime import datetime
import tkinter as tk
import numpy as np
import win32print
import win32api
import traceback
import threading
import warnings
import smtplib
import json
import time
import fitz
import glob
import re
import os

warnings.simplefilter("ignore", UserWarning)

with open(os.path.join(os.path.dirname(__file__), "config.json")) as f:
    config = json.load(f)

class RectangulatorHandler:

    def __init__(self, root, fig, ax):
        self.queue = []
        self.invoice = True
        self.should_print = True
        self.should_save = True
        self.hit_submit = False
        self.root = root
        self.fig = fig
        self.ax = ax
        self.done_var = tk.IntVar()
        self.config = config


    def refresh_config(self):
        with open(os.path.join(os.path.dirname(__file__), "config.json"), "r") as f:
            self.config = json.load(f)

    def rectangulate(self, filename, filepath, root, template_folder, testing=False):  # Add a file to the queue, or process it immediately if template exists
        # If in away mode, just print and move to away label
        if root.AWAY_MODE and not testing:
            self.print_invoice(filepath)
            return

        self.log(f"Template required for {filename}", display=True)

        try:
            # Tries to return list of [filename, should_print] or [None, False], or special flags
            if not testing:
                self.send_email("Must create template", root)  # email me

            # Use rectangulator
            if testing == "scanner":
                rectangulator, text_box = self.open_rectangulator(filepath, template_folder, root, scanner=True) 
            else:
                rectangulator, text_box = self.open_rectangulator(filepath, template_folder, root)

            # Process the rectangulator output
            if testing == "test":  # if using text inbox
                return ["test_email"]
            if not rectangulator and not text_box:  # if the window was closed
                return []
            if not self.invoice:  #  if the user clicked the "Not An Invoice" button
                self.invoice = True
                if text_box.text:
                    filename = text_box.text
                return ["not_invoice", [os.path.join(os.path.dirname(filepath), f"{filename}.pdf"), self.should_print, self.should_save]]
            filename = rectangulator.rename_pdf()
            if filename:  # if the user dragged a rectangle
                return [filename, self.should_print]
            elif text_box.text:  # if the user entered a filename
                return [os.path.join(os.path.dirname(filepath), f"{text_box.text}.pdf"), self.should_print]
        except Exception as e:
            self.log(f"An error occurred while rectangulating: {str(e)}", tag="red", display=True)
        return [None, False]

    def check_templates(self, pdf_path, template_folder, root):  # Check if a template exists for the invoice
        for file in glob.glob(rf"{template_folder}\*.txt"):
            try:
                with open(file, "r") as f:
                    while True:
                        # Get the invoice name, date, and number from the template
                        invoice_name = f.readline().split("?")
                        invoice_date = f.readline().split("?")
                        invoice_num = f.readline().split("?")

                        if not invoice_name or not invoice_date or not invoice_num:
                            break

                        # Get the company name from the invoice
                        identifier = self.get_text_in_rect(Rectangle((invoice_name[1], invoice_name[2]), invoice_name[3], invoice_name[4]), pdf_path)

                        # If company name on invoice matches name on template, use that template
                        if identifier.strip() and invoice_name[0] == identifier:
                            self.log(f"Used template {file} for {identifier}")
                            # Get the invoice date and number from the invoice
                            invoice_date = self.get_text_in_rect(Rectangle((invoice_date[1], invoice_date[2]), invoice_date[3], invoice_date[4]), pdf_path)
                            invoice_num = self.get_text_in_rect(Rectangle((invoice_num[1], invoice_num[2]), invoice_num[3], invoice_num[4]), pdf_path)
                            # Clean the invoice date
                            invoice_date = self.clean_date(invoice_date)
                            return [rf"{os.path.dirname(pdf_path)}\{invoice_date}_{invoice_num}.pdf"]
            except Exception as e:
                pass

    def open_rectangulator(self, pdf_path, template_folder, root, scanner=False):  # Setup the page for the Rectangulator and return the Rectangulator and textbox
        # Reset flags
        self.should_print = True 
        self.should_save = True
        self.hit_submit = False
        self.invoice = True

        # Don't print by default when using scanner
        if scanner:
            self.should_print = False  

        self.done_var.set(0)  # reset done variable
        doc = fitz.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap()
        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        doc.close()

        self.ax.clear()
        self.ax.imshow(img_array)
        self.ax.axis("off")
        self.fig.canvas.draw()

        # Create a Not An Invoice button
        not_inv_button_ax = self.fig.add_axes([0.64, 0.005, 0.18, 0.075])
        not_inv_button = Button(not_inv_button_ax, "Not An Invoice")

        def not_invoice(event):  # If the user clicks the "Not Invoice" button
            try:
                self.invoice = False
                self.ax.clear()
                self.ax.axis("off")
                self.fig.canvas.draw_idle()
                self.done_var.set(1)
            except Exception as e:
                self.log(f"Error in not_invoice: {str(e)} \n{traceback.format_exc()}")
                self.done_var.set(1) 

        not_inv_button.on_clicked(not_invoice)

        # Create a checkbox for if it should be printed
        print_checkbox_ax = self.fig.add_axes([0.86, 0.03, 0.03, 0.03])
        print_checkbox = CheckButtons(print_checkbox_ax, [""], [self.should_print])
        def print_callback(label):
            self.should_print = not self.should_print
        print_checkbox.on_clicked(print_callback)
        # Set the checkbox to be a square and centered
        for i, line in enumerate(print_checkbox.lines):
            rect = print_checkbox.rectangles[i]
            rect.set_width(1)
            rect.set_height(1)
            rect.set_edgecolor("none")
            # Calculate the center of the rectangle
            center_x = rect.get_width() / 2
            center_y = rect.get_height() / 2
            # Update the line positions to be centered
            line[0].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[1].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[0].set_ydata([center_y - rect.get_height() / 4, center_y + rect.get_height() / 4])
            line[1].set_ydata([center_y + rect.get_height() / 4, center_y - rect.get_height() / 4])
        print_label = self.fig.text(0.853, 0.075, "Print?", fontsize=9)

        # Create a checkbox for if it should be saved
        save_checkbox_ax = self.fig.add_axes([0.93, 0.03, 0.03, 0.03])
        save_checkbox = CheckButtons(save_checkbox_ax, [""], [True])
        def save_callback(label):
            self.should_save = not self.should_save
        save_checkbox.on_clicked(save_callback)
        # Set the checkbox to be a square and centered
        for i, line in enumerate(save_checkbox.lines):
            rect = save_checkbox.rectangles[i]
            rect.set_width(1)
            rect.set_height(1)
            rect.set_edgecolor("none")
            # Calculate the center of the rectangle
            center_x = rect.get_width() / 2
            center_y = rect.get_height() / 2
            # Update the line positions to be centered
            line[0].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[1].set_xdata([center_x - rect.get_width() / 4, center_x + rect.get_width() / 4])
            line[0].set_ydata([center_y - rect.get_height() / 4, center_y + rect.get_height() / 4])
            line[1].set_ydata([center_y + rect.get_height() / 4, center_y - rect.get_height() / 4])
        save_label = self.fig.text(0.92, 0.075, "Save?", fontsize=9)

        # Filename text box and submit button
        text_box_ax = self.fig.add_axes([0.1, 0.005, 0.35, 0.075])
        text_box = TextBox(text_box_ax, label="", initial="")

        def on_text_submit(text):
            if self.hit_submit:
                return
            self.hit_submit = True

            def answer_thread():
                filename_is_correct = self.create_alert(f"Is '{text_box.text}' the correct filename?")
                self.hit_submit = False
                if filename_is_correct:
                    self.ax.clear()
                    self.ax.axis("off")
                    self.fig.canvas.draw_idle()
                    self.done_var.set(1)  # signal that the user is done

            # run in a separate thread to avoid blocking the main thread
            threading.Thread(target=answer_thread).start()  

        submit_button_ax = self.fig.add_axes([0.45, 0.005, 0.15, 0.075])
        submit_button = Button(submit_button_ax, "Submit")
        submit_button.on_clicked(on_text_submit)

        # Create text labels for instructions and text box
        text_box_label = self.fig.text(
            0.1,
            0.089,
            "Enter Filename Manually (mm-dd-yy_invoice#)",
            fontsize=10)
        instruction_label = self.fig.text(
            0.1,
            0.975,
            "- Left Click to Draw boxes around Company Name, Date, and Invoice (in that order)",
            fontsize=10)
        instruction_label_2 = self.fig.text(
            0.1,
            0.95,
            "- Company Name can be any piece of text unique to that vendor",
            fontsize=10)
        instruction_label_3 = self.fig.text(
            0.1,
            0.925,
            "- Right Click to verify and save, Middle Click to Pan, Scroll to Zoom",
            fontsize=10)

        self.fig.canvas.draw_idle()
        rectangulator = Rectangulator(self.ax, self.fig, pdf_path, template_folder, self)
        root.root.wait_variable(self.done_var)  # wait for the user to finish

        # Remove the text labels and rectangles
        for ax in [text_box_ax, not_inv_button_ax, print_checkbox_ax, save_checkbox_ax, submit_button_ax]:
            self.fig.delaxes(ax)
        for label in [text_box_label, instruction_label, instruction_label_2, instruction_label_3, print_label, save_label]:
            label.remove()
        for cid in rectangulator.cids:
            self.fig.canvas.mpl_disconnect(cid)
        rectangulator.cids.clear()
        self.ax.clear()
        self.ax.axis("off")
        self.fig.canvas.draw_idle()

        return rectangulator, text_box

    def print_invoice(self, filepath):  # Printer
        try:
            # Get default printer and print
            p = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, "print", filepath, None, ".", 0)
            self.log(f"Printed {os.path.basename(filepath)} completed successfully.", tag="lgreen", display=True)
            return True
        except Exception as e:
            self.log(f"Printing failed for {filepath}: {str(e)}", tag="red", send_email=True, display=True)
            return False

    def log(self, *args, tag="purple", send_email=False, display=False):  # Log messages to a file and optionally send an email
        message = " ".join([str(arg) for arg in args])
        if display:
            self.root.log(message, tag=tag, send_email=send_email)
        else:
            print(f"-{message}")
        if send_email:
            self.send_email(message, self.root)

        with open(self.config["LOG_FILE"], "a") as file:
            file.write(message + "\n")

    def sanitize_filename(self, filename):  # Remove invalid characters from the filename
        sanitized_filename = re.sub(r"[^\w_. -]", "", filename.replace("/", "-"))
        return sanitized_filename.strip()

    def get_text_in_rect(self, rect, pdf_path):  # Get the text within a specified rectangle in a PDF document
        doc = None
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
                word_x, word_y, _, _ = word[:4]  # get word coordinates
                if (x <= float(word_x) <= x + width) and (y <= float(word_y) <= y + height):
                    extracted_text += word[4] + " "  # append the word to extracted text
            extracted_text = extracted_text.strip()  # remove leading/trailing spaces

            return extracted_text
        except Exception as e:
            self.log(f"An error occurred while processing the PDF: {str(e)} {traceback.format_exc()}")
            return ""
        finally:
            if doc:
                doc.close()

    def check_date_outlier(self, invoice_name, invoice_date):  # Check if the date is an outlier and correct it
        calendar = {
            "Jan": "01",
            "Feb": "02",
            "Mar": "03",
            "Apr": "04",
            "May": "05",
            "Jun": "06",
            "Jul": "07",
            "Aug": "08",
            "Sep": "09",
            "Oct": "10",
            "Nov": "11",
            "Dec": "12"
        }
        calendar2 = {
            "January": "01",
            "February": "02",
            "March": "03",
            "April": "04",
            "May": "05",
            "June": "06",
            "July": "07",
            "Aug": "08",
            "September": "09",
            "October": "10",
            "November": "11",
            "December": "12"
        }

        try:
            # Uses format "DD-Month-YY"
            invoice_date = invoice_date.split("-")
            day = invoice_date[0]
            invoice_date[0] = calendar[invoice_date[1]]
            invoice_date[1] = day
            invoice_date = "-".join(invoice_date)
            if datetime.striptime(invoice_date, "%m-%d-%y"):
                return invoice_date
        except:
            pass
        
        try:
            # Uses format "Month DD, YYYY"
            invoice_date = invoice_date.replace(",", "").split(" ")
            invoice_date[0] = calendar2[invoice_date[0]]
            invoice_date = "/".join(invoice_date)
        except:
            pass

        try:
            # Uses format "Mon DD, YYYY"
            invoice_date = invoice_date.replace(",", "").split(" ")
            invoice_date[0] = calendar[invoice_date[0]]
            invoice_date = "/".join(invoice_date)
        except:
            pass

        return self.clean_date(invoice_date.strip())

    def clean_date(self, invoice_date):  # Clean the date to be in the format "MM-DD-YY"
        date_patterns = [
            "%B %d, %Y",
            "%b %d, %Y",
            "%B %d, %y",
            "%b %d, %y",
            "%d-%B-%Y",
            "%d-%b-%Y",
            "%d-%B-%y",
            "%d-%b-%y",
            "%m-%d-%Y",
            "%m-%d-%y",
            "%b %d %Y",
            "%B %d %Y",
            "%b %d %y",
            "%B %d %y",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%Y-%m-%d",
            "%y-%m-%d",
            "%Y/%m/%d",
            "%y/%m/%d"
        ]
        invoice_date = str(invoice_date).replace("/", "-")
        for pattern in date_patterns:
            try:
                dt = datetime.strptime(invoice_date, pattern)
                return dt.strftime("%m-%d-%y")
            except ValueError:
                continue
        self.log(f"Could not convert {invoice_date} to date")
        return invoice_date

    def send_email(self, body, root):  # Sends email to me
        if root:
            sender_email = f"{root.username}.sndex@gmail.com"
            password = root.password
            if root.TESTING:
                return
        else:
            sender_email = f"{self.config['ACP_USER']}.sndex@gmail.com"
            password = self.config['ACP_PASS']

        try:
            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["Subject"] = "Alert"
            message["From"] = sender_email
            message["To"] = self.config["RECEIVER_EMAIL"]
            message.attach(MIMEText(body, "plain"))

            # Send the email
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, password)
                server.sendmail(sender_email, self.config["RECEIVER_EMAIL"], message.as_string())
            self.log(f"Email sent successfully: {body}")
        except Exception as e:
            self.log(f"Error sending email {body} - {str(e)}")

    def create_alert(self, message, numbered_buttons=0):  # Create an alert window for user input
        try:
            parent = self.root.alert_container
            panel = AlertWindow(parent, message, numbered_buttons)
            panel.pack(fill=tk.BOTH, expand=True)
            parent.lift()
            panel.grab_set()
            panel.focus_set()
            answer = panel.get_answer()  # wait for user input
            panel.destroy()  # destroy the alert window
            parent.lower()  # lower the alert container
            return answer
        except Exception as e:
            self.log(f"Error creating alert: {str(e)} \n{traceback.format_exc()}")
            return False

class Rectangulator:

    def __init__(self, ax, fig, pdf_path, template_folder, rectangulator_handler):
        self.rectangulator_handler = rectangulator_handler
        self.pdf_path = pdf_path
        self.template_folder = template_folder
        self.fig = fig
        self.ax = ax

        self.rectangles = []  # contains rectangle objects
        self.coordinates = []  # contains coordinates of rectangle objects
        self.correcting_rect_index = None  # used when redrawing specific rectangle
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
        self.cids = []
        cid1 = self.ax.figure.canvas.mpl_connect("button_press_event", self.on_button_press)
        cid2 = self.ax.figure.canvas.mpl_connect("button_release_event", self.on_button_release)
        cid3 = self.ax.figure.canvas.mpl_connect("motion_notify_event", self.on_move)
        cid4 = self.ax.figure.canvas.mpl_connect("scroll_event", self.on_scroll)
        cid5 = self.ax.figure.canvas.mpl_connect("key_press_event", self.on_key_press)
        self.cids.extend([cid1, cid2, cid3, cid4, cid5])

    def rename_pdf(self):  # Rename the PDF based on extracted text from rectangles
        try:
            if len(self.rectangles) == 3 and all(
                    self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path) for rect in self.rectangles):
                self.rectangulator_handler.log(f"Creating new template")
                self.save_template()

                # Rename the PDF based on extracted text from rectangles in format "MM-DD-YY_INVOICE_NUMBER"
                extracted_texts = [
                    self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path) for rect in self.rectangles
                    if self.rectangulator_handler.get_text_in_rect(rect, self.pdf_path)
                ]
                extracted_texts[1] = self.rectangulator_handler.clean_date(extracted_texts[1])  # fix date
                extracted_text_combined = "_".join(extracted_texts[1:])  # combine date and invoice number
                sanitized_extracted_text = self.rectangulator_handler.sanitize_filename(extracted_text_combined)  # sanitize the text
                new_filepath = os.path.join(os.path.dirname(self.pdf_path), f"{sanitized_extracted_text}.pdf")  # combine the path and filename
                return new_filepath
        except RecursionError:
            self.rectangulator_handler.log("Window closed please try again")
            return None
        except Exception as e:
            self.rectangulator_handler.log(traceback.format_exc())
            self.rectangulator_handler.log(
                f"Error occurred with download, {str(e)}")
            return None

    def save_template(self):  # Save the template to a text file
        # Save the template to a text file as Company Name?x?y?width?height and so on
        filename = rf"{self.template_folder}\{self.rectangulator_handler.sanitize_filename(self.rectangulator_handler.get_text_in_rect(self.rectangles[0], self.pdf_path))}{str(int(time.time()))}.txt"
        # Check if the file already exists or filename is empty
        if filename == "":
            self.rectangulator_handler.log(f"Filename empty, not creating template")
            return
        with open(filename, "a") as file:
            for i, coord in enumerate(self.coordinates):
                x, y, width, height = coord
                rect_text = self.rectangulator_handler.get_text_in_rect(self.rectangles[i], self.pdf_path)
                text = f"{rect_text}?{x}?{y}?{width}?{height}\n"
                if i == 0:  # need to sanitize company name
                    file.write(f"{self.rectangulator_handler.sanitize_filename(rect_text)}?{x}?{y}?{width}?{height}\n")
                else:
                    file.write(text)
        self.rectangulator_handler.log(f"Created invoice template {filename}")

    def on_key_press(self, event):  # Handle key press events
        if event.key == "escape":  # reset zoom and position
            self.ax.set_xlim(self.initial_xlim)
            self.ax.set_ylim(self.initial_ylim)
            self.pan_start = None
            self.ax.figure.canvas.draw()

    def on_button_press(
            self, event):  # Handle left and right mouse button press events
        if event.button == 1:  # left mouse button, draw rectangles
            # Ignore if the mouse click is outside the plot area
            if event.xdata is None or event.ydata is None:
                return

            # Start drawing a rectangle
            self.start_x = event.xdata
            self.start_y = event.ydata
            self.rect = Rectangle((self.start_x, self.start_y), 0, 0, edgecolor="red", linewidth=2, fill=False)
            self.ax.add_patch(self.rect)
            self.ax.figure.canvas.draw()

        elif event.button == 2:  # middle mouse button, pan
            self.pan_start = (event.x, event.y)

        elif event.button == 3:  # right mouse button, save rectangles
            # Check if 3 rectangles have been drawn
            if len(self.rectangles) != 3:
                self.rectangulator_handler.log(
                    "Please draw exactly three rectangles", display=True)
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
                text_is_correct = self.rectangulator_handler.create_alert(
                    f"Does the following text match what you selected?\n\n{extracted_text}",
                    numbered_buttons=3)

                # If user says yes, close the window, otherwise reset the rectangles
                if isinstance(text_is_correct,int) and not isinstance(text_is_correct, bool):
                    self.correcting_rect_index = text_is_correct
                    self.rectangulator_handler.log(f"Please reselect {headers[self.correcting_rect_index]}")
                    self.reset_rectangles(specific_rect=self.correcting_rect_index)
                elif text_is_correct:
                    self.ax.clear()
                    self.ax.axis("off")
                    self.fig.canvas.draw_idle()
                    self.rectangulator_handler.done_var.set(1)
                else:
                    self.rectangulator_handler.log("Please reselect rectangles", display=True)
                    self.reset_rectangles()

            input_thread = threading.Thread(target=handle_user_input)
            input_thread.start()

    def on_button_release(self, event):  # Handle key release events
        if event.button == 1 and self.rect:  # left mouse button, save rectangle
            self.start_x = None
            self.start_y = None

            # Append rectangle and coordinated to list
            self.rectangles.append(self.rect)
            self.coordinates.append((self.rect.get_x(), self.rect.get_y(), self.rect.get_width(),
                                     self.rect.get_height()))  # store coordinates
            self.rect = None
            self.ax.figure.canvas.draw()

        elif event.button == 2:  # middle mouse button, stop panning
            self.pan_start = None
            self.prev_x = None
            self.prev_y = None

    def on_move(self, event):  # Continuously update drawn rectangles
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

    def on_scroll(self, event):  # Zoom in and out
        if event.button == "down":  # out
            self.zoom(event.xdata, event.ydata, 1 / self.zoom_factor)
        elif event.button == "up":  # in
            self.zoom(event.xdata, event.ydata, self.zoom_factor)

    def reset_rectangles(self, specific_rect=None):  # Reset the current rectangles
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

    def zoom(self, x, y, zoom_factor):  # Zoom in and out with scroll wheel
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        if xlim is None or ylim is None:
            return

        # Calculate new limits
        if x is None:
            x = np.mean(xlim)
        if y is None:
            y = np.mean(ylim)
        new_xlim = ((xlim[0] - x) / zoom_factor) + x, (
            (xlim[1] - x) / zoom_factor) + x
        new_ylim = ((ylim[0] - y) / zoom_factor) + y, (
            (ylim[1] - y) / zoom_factor) + y

        self.ax.set_xlim(new_xlim)
        self.ax.set_ylim(new_ylim)
        self.ax.figure.canvas.draw()