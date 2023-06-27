from email.mime.multipart import MIMEMultipart
from matplotlib.patches import Rectangle
from matplotlib.widgets import Button
from email.mime.text import MIMEText
from datetime import datetime
import matplotlib.pyplot as plt
import tkinter as tk
import numpy as np
import pytesseract
import traceback
import threading
import keyboard
import smtplib
import config
import math
import fitz
import glob
import re
import os

pytesseract.pytesseract.tesseract_cmd = config.PYTESSERACT_PATH
log_file = result = root = None


class Rectangulator:

    def __init__(self, ax, fig, pdf_path, template_folder):
        global log_file
       
        # VARIABLES
        self.pdf_path = pdf_path
        self.template_folder = template_folder
        self.fig = fig
        self.ax = ax

        self.rectangles = [] #contains rectangle objects
        self.coordinates = [] #contains coordinates of rectangle objects
        self.start_x = None
        self.start_y = None
        self.rect = None 
        self.zoom_factor = 1.2

        # Connect the event handlers to the canvas
        self.ax.figure.canvas.mpl_connect('button_press_event', self.on_button_press)
        self.ax.figure.canvas.mpl_connect('button_release_event', self.on_button_release)
        self.ax.figure.canvas.mpl_connect('motion_notify_event', self.on_move)
        self.ax.figure.canvas.mpl_connect('scroll_event', self.on_scroll)

    def on_button_press(self, event): # Handle left and right mouse button press events
        if event.button == 1:  # Left mouse button
            # Ignore if the mouse click is outside the plot area
            if event.xdata is None or event.ydata is None:
                return 
            
            self.start_x = event.xdata
            self.start_y = event.ydata
            self.rect = Rectangle((self.start_x, self.start_y), 0, 0, edgecolor='red', linewidth=2, fill=False)
            self.ax.add_patch(self.rect)
            self.ax.figure.canvas.draw()
        elif event.button == 3:  # Right mouse button
            # Check if 3 rectangles have been drawn
            if len(self.rectangles) != 3:
                log("Please draw exactly three rectangles")
                self.reset_rectangles()
                return
            
            # Show extracted text for verification
            log("Extracted Text:")
            headers = ["--- Company Name: ", "--- Invoice Date: ", "--- Invoice Number: "]
            for i, rect in enumerate(self.rectangles):
                extracted_text = headers[i] + get_text_in_rect(rect, self.pdf_path)
                log(extracted_text)

            # Ask for verifictaion
            def handle_user_input():
                while True:
                    print("Does the following text match what you selected? (press y/n): ")
                    choice = keyboard.read_event()
                    if choice.event_type == "down":
                        if choice.name.lower() == "y":
                            plt.close(self.fig)
                            break
                        elif choice.name.lower() == "n":
                            log("Please reselect rectangles")
                            self.reset_rectangles()
                            break
                        else:
                            log("Invalid choice. Please enter 'y' or 'n'.")

            input_thread = threading.Thread(target=handle_user_input)
            input_thread.start()
         
    def on_button_release(self, event): # Save rectangle coordinates
        if event.button != 1 or self.rect is None:
            return
        self.start_x = None
        self.start_y = None

        # Append rectangle and coordinated to list
        self.rectangles.append(self.rect)
        self.coordinates.append((self.rect.get_x(), self.rect.get_y(), self.rect.get_width(), self.rect.get_height()))  # Store coordinates
        self.rect = None
        self.ax.figure.canvas.draw()

    def on_move(self, event): # Continuously update drawn rectangle
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
        if event.button == 'down':
            self.zoom(event.xdata, event.ydata, 1 / self.zoom_factor)
        elif event.button == 'up':
            self.zoom(event.xdata, event.ydata, self.zoom_factor)

    def reset_rectangles(self): # Reset the current rectangles
        for rect in self.rectangles:
            rect.remove()
        self.rectangles = []
        self.coordinates = []
        self.rect = None  
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

    def rename_pdf(self):
        try:
            if len(self.rectangles) == 3:
                self.save_template()
            
                # Rename the PDF based on extracted text from rectangles in format "MM-DD-YY_INVOICE_NUMBER"
                extracted_texts = [get_text_in_rect(rect, self.pdf_path) for rect in self.rectangles if get_text_in_rect(rect, self.pdf_path)]
                extracted_texts[1] = check_outlier(extracted_texts[0], extracted_texts[1])
                extracted_text_combined = '_'.join(extracted_texts[1:])
                sanitized_extracted_text = sanitize_filename(extracted_text_combined)
                new_filename = os.path.join(os.path.dirname(self.pdf_path), f"{sanitized_extracted_text}.pdf")
                return new_filename 
        except RecursionError:
            log("Window closed please try again")
            return None
        except Exception as e:
            print(traceback.format_exc())
            log(f"Error occurred with download, {str(e)}")
            return None
        
    def save_template(self):
        # Save the template to a text file as Company Name?x?y?width?height and so on
        for i, coord in enumerate(self.coordinates):
            x, y, width, height = coord
            rect_text = get_text_in_rect(self.rectangles[i], self.pdf_path)
            text = f"{rect_text}?{x}?{y}?{width}?{height}\n"
            filename = rf"{self.template_folder}\{sanitize_filename(get_text_in_rect(self.rectangles[0], self.pdf_path))}.txt"
            with open(filename, 'a') as file:
                if i == 0:
                    file.write(f"{sanitize_filename(rect_text)}?{x}?{y}?{width}?{height}\n")
                else:
                    file.write(text)
        log(f"Created vendor invoice template {get_text_in_rect(self.rectangles[0], self.pdf_path)}")


def sanitize_filename(filename): # Remove invalid characters from the filename
    sanitized_filename = re.sub(r'[^\w_. -]', '', filename.replace('/', '-'))
    return sanitized_filename.strip()

def get_text_in_rect(rect, pdf_path):
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
            word_x, word_y, _, _ = word[:4] #get word coordinates
            if x <= float(word_x) <= x + width and y <= float(word_y) <= y + height:
                extracted_text += word[4] + " " #append the word to extracted text
        extracted_text = extracted_text.strip() #remove leading/trailing spaces

        # If no text was extracted, try OCR using pytesseract
        page_width, page_height = page.rect.width, page.rect.height
        x, y, width, height = math.ceil(float(x)), math.ceil(float(y)), math.ceil(float(width)), math.ceil(float(height))
        if extracted_text == "" and x >= 0 and y >= 0 and x + width <= page_width and y + height <= page_height: 
            pix = page.get_pixmap()
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            # Extract the region of interest and convert to grayscale then use OCR
            img_region = img_array[y:y + height, x:x + width, :]
            img_gray = np.mean(img_region, axis=2).astype(np.uint8)
            extracted_text = pytesseract.image_to_string(img_gray)

        # If OCR fails, set extracted text to ""
        if extracted_text == None:
            extracted_text = ""

        return extracted_text
    except Exception as e:
        log(f"An error occurred while processing the PDF: {str(e)} {traceback.format_exc()}")
        return ""
    finally:
        doc.close()

def check_outlier(invoice_name, invoice_date):
    calendar = {"Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04", "May": "05", "Jun": "06", "Jul": "07", "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", 'Dec': "12"}
    if invoice_name == "BUZZI UNICEM USA - Cement":
        # Uses format "DD-Month-YY"
        invoice_date = invoice_date.split("-")
        day = invoice_date[0]
        invoice_date[0] = calendar[invoice_date[1]]
        invoice_date[1] = day
        return "-".join(invoice_date)
    elif invoice_name in {"ADP SCREENING  SELECTION SERVICES", "Muka Development Group Llc"}:
        # Uses format "Month DD, YYYY"
        invoice_date = invoice_date.replace(",", "").split(" ")
        invoice_date[0] = calendar[invoice_date[0]]
        invoice_date = "/".join(invoice_date)
    
    return clean_date(invoice_date)

def clean_date(invoice_date):
    try:
        date = datetime.strptime(invoice_date, "%m/%d/%y")
        return date.strftime("%m/%d/%y")
    except ValueError:
        try:
            date = datetime.strptime(invoice_date, "%m/%d/%Y")
            return date.strftime("%m/%d/%y")
        except ValueError:
            return invoice_date
    
def log(*args):
    global log_file
    global root
    with open(log_file, "a") as file:
        message = "#RECTANGULATOR# " + ' '.join([str(arg) for arg in args]) 
        print(message)
        file.write('\n'.join([str(arg) for arg in args]))

def send_email():
    try:
        sender_email = f"{config.ACP_USER}{config.ADDRESS}"
        reciever_email = config.RECIEVER_EMAIL

        # Create a multipart message
        message = MIMEMultipart()
        message["Subject"] = "Must create template"
        message["From"] = sender_email
        message["To"] = reciever_email
        message.attach(MIMEText("Must create template", "plain"))

        # Send the email
        with smtplib.SMTP(config.SMTP_SERVER, 587) as server:
            server.starttls()
            server.login(sender_email, config.ACP_PASS)
            server.sendmail(sender_email, reciever_email, message.as_string())
            log(f"Template request sent from {sender_email} to {reciever_email}")
    except Exception as e:
            log(f"Error sending email from {sender_email} - {str(e)}")

def not_invoice(event):
    # If the user clicks the "Not Invoice" button, set the global result to False and close the window
    global result
    result = True
    plt.close()

def main(pdf_path, template_folder, log_file_arg, root_arg):
    global log_file
    global root
    log_file = log_file_arg
    root = root_arg

    # Iterate through invoice templates and check for one that matches the invoice
    for file in glob.glob(rf"{template_folder}\*.txt"):
        try:
            with open(file, "r") as f:
                # Get the invoice name, date, and number from the template
                invoice_name = f.readline().split("?")
                invoice_date = f.readline().split("?")
                invoice_num = f.readline().split("?")
                # Get the company name from the invoice
                identifier = sanitize_filename(get_text_in_rect(Rectangle((invoice_name[1], invoice_name[2]), invoice_name[3], invoice_name[4]), pdf_path))

                # If company name on invoice matches name on template, use that template
                if invoice_name[0] == identifier:
                    # Get the invoice date and number from the invoice
                    invoice_date = get_text_in_rect(Rectangle((invoice_date[1], invoice_date[2]), invoice_date[3], invoice_date[4]), pdf_path)
                    invoice_num = get_text_in_rect(Rectangle((invoice_num[1], invoice_num[2]), invoice_num[3], invoice_num[4]), pdf_path)
                    # Clean the invoice date
                    invoice_date = check_outlier(invoice_name[0], invoice_date).replace("/", "-")

                    return rf"{os.path.dirname(pdf_path)}\{invoice_date}_{invoice_num}.pdf"
        except Exception as e:
            pass

    # If no template exists, make one
    try:
        send_email() #email me

        # Draw rectangles on the PDF image for annotation
        doc = fitz.open(pdf_path)
        page = doc[0]

        # Convert the page to a NumPy array for plotting
        pix = page.get_pixmap()
        global img_array
        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)

        # Create a matplotlib figure and axis for displaying the image
        fig, ax = plt.subplots(figsize=(8, 8))
        plt.subplots_adjust(bottom=0.2)

        # Create a button and specify its position and label
        button_ax = plt.axes([0.7, 0.05, 0.1, 0.075])  #[left, bottom, width, height]
        button = Button(button_ax, 'Not An Invoice')

        # Assign the button_pressed function as the callback when the button is clicked
        button.on_clicked(not_invoice)
        ax.imshow(img_array)

        # Create an instance of DraggableRectangle and bind it to the axis
        draggable_rect = Rectangulator(ax, fig, pdf_path, template_folder)

        # Show the plot with the PDF image and rectangles
        plt.show()

        # If the user clicked the "Not An Invoice" button
        if result:
            return "not_invoice"

        # Return the new filename created by renaming the PDF
        return draggable_rect.rename_pdf()
    except Exception as e:
        log(f"An error occurred while drawing rectangles: {str(e)}")

    return None