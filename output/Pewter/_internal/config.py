from math import ceil
import os
LOG_FILE = 'C:\\Users\\jmwesthoff\\OneDrive - atlanticconcrete.com\\Documents\\Scripts\\Pewter\\Log.txt'
INVOICE_FOLDER = 'S:\\Titan_DM\\Titan_Inbox\\AP_Emailed_Invoices'
TEMPLATE_FOLDER = 'S:\\Titan_DM\\Titan_Inbox\\AP_Emailed_Invoices\\Templates'
STATEMENT_FOLDER = "S:\\Titan_DM\\AP\\Vendor Statements"
TEST_INVOICE = "S:\\Titan_DM\\Titan_Inbox\\AP_Emailed_Invoices\\Test Invoices\\Test Invoice.pdf"
TEST_INVOICE_FOLDER = 'S:\\Titan_DM\\Titan_Inbox\\AP_Emailed_Invoices\\Test Invoices'
TEST_TEMPLATE_FOLDER = 'S:\\Titan_DM\\Titan_Inbox\\AP_Emailed_Invoices\\Test Invoices'
ACP_USER = "concrete"
ACP_PASS = "wggfkcyusmyklrgc"
APC_USER = 'precast'
APC_PASS = 'yolhphmgrbzqqxcn'
IMAP_SERVER = "imap.gmail.com"
SMTP_SERVER = "smtp.gmail.com"
RECEIVER_EMAIL = 'jmwesthoff@atlanticconcrete.com'
TRUSTED_ADDRESS = "atlanticconcrete.com"
SCANNER_EMAIL = 'jmwesthoff@atlanticconcrete.com'
ADDRESS = ".sndex@gmail.com"
INBOX_CYCLE_TIME = 5
RECONNECT_TIME = 3600
RECONNECT_CYCLE_COUNT = ceil(RECONNECT_TIME / INBOX_CYCLE_TIME)
LAST_CRASH_DATE = '2025-08-01'