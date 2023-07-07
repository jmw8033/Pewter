# Pewter
Emailed invoice file handler.
In collaboration with the Rectangulator and Loginulator.

## What's it do?
It downloads invoices from an email, correctly names them, and prints them.

## Why?
Because downloading things is a pain.

## How's it do it?
It first connects to the emails (I use two emails for super secret reasons), then waits for an email, and if it gets one from a trusted email, it downloads the attachments.
Then using the Rectangulator (trademarked) it saves a template containing where in the PDF the name of the company, invoice date, and invoice number are saved, so that
when an invoice is recieved from that same company, it can correctly name it on its own (we name invoices in the format mm-dd-yy_invoice#). It also prints it to the default printer.

## Configuration
The program is currently really only built to be used for my specific purpose, but if you wanted to use it for your own project, the only code you would really have to modify
is the main method in the EmailProcessor class and the Rectangulator to fit your invoice name format (and probably some other things since I didn't think this project would get this far and it's a bit messy). 

**You would also have to modify the config.py file:**
* LOG_FILE: Path to the log text file for storing application logs.
* TEMPLATE_FOLDER: Path to the folder containing invoice templates.
* INVOICE_FOLDER: Path to the folder where downloaded invoices will be stored.
* ACP_USER, ACP_PASS, APC_USER, APC_PASS: Username and passwords for email accounts. **These are purely unique to me. If you're just handling one email you would just have one set of username and password.**
* IMAP_SERVER: IMAP server address (ex. imap.gmail.com).
* SMTP_SERVER: SMTP server address (ex. smtp.gmail.com).
* RECIEVER_EMAIL: Email address to receive error alerts.
* TRUSTED_ADDRESS: Trusted email address for invoice senders.
* ADDRESS: Email address domain (ex. @gmail.com).

## Contributing
Contributions to this project are welcome. If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.
