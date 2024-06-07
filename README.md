# Pewter
Emailed invoice file handler.
In collaboration with the Rectangulator and Loginulator.

## What's it do?
It downloads invoices from an email, correctly names them, and prints them.

## Why?
Because downloading things is a pain.

## How's it do it?
It first connects to the emails (in our case we have two), then waits for an email, and if it gets one from a trusted email, it downloads the attachments.
Then using the Rectangulator (trademarked), the user saves a template containing where in the PDF the name of the company, invoice date, and invoice number are saved, so that
when an invoice is recieved from that same company, it can correctly name it on its own (in the format mm-dd-yy_invoice#). It also prints it to the default printer.

## Configuration
The program is now significantly more accessible as each email is now opened in its own window. You can now easily set up as many emails as you want and all you need to change is the config and the Rectangulator if 
you need a different invoice name format.

**Current config.py file setup:**
* LOG_FILE: Path to the log text file for storing application logs.
* INVOICE_FOLDER: Path to the folder where downloaded invoices will be stored.
* TEMPLATE_FOLDER: Path to the folder containing invoice templates.
* TEST_INVOICE: Path to pdf of test invoice used for rectangulator testing.
* TEST_INVOICE_FOLDER: Path to the folder where downloaded invoices will be store when testing is enabled.
* TEST_TEMPLATE_FOLDER: Path to the folder containing invoice templates when testing is enabled.
* ACP_USER, ACP_PASS, APC_USER, APC_PASS: Username and passwords for email accounts. **These are purely unique to me. If you're just handling one email you would just have one set of username and password.**
* IMAP_SERVER: IMAP server address (ex. imap.gmail.com).
* SMTP_SERVER: SMTP server address (ex. smtp.gmail.com).
* RECIEVER_EMAIL: Email address to receive error alerts.
* TRUSTED_ADDRESS: Trusted email address for invoice senders.
* ADDRESS: Email address domain (ex. @gmail.com).
* PYTESSERACT_PATH: Path to pytesseract
* CHROMEDRIVER_PATH: Path to chromedriver
* WAIT_TIME = Cycle time for searching inbox (in seconds).
* RECONNECT_CYCLE_COUNT = Number of cycles until it reconnects to email (ex. 3600 / WAIT_TIME would be one hour).
* RECTANGULATOR_TIMEOUT = Timeout time for rectangulator (in milliseconds).

## Contributing
Contributions to this project are welcome. If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.
