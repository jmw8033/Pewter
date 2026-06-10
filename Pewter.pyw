from EmailProcessor import EmailProcessor
import json, os, keyring

def get_password(username):
    password = keyring.get_password("PewterInvoiceProcessor", username)
    if password is None:
        print(f"No stored password for '{username}'. Store one with:")
        print(f"  keyring.set_password('PewterInvoiceProcessor', '{username}', '<app password>')")
        exit(1)
    return password

with open(os.path.join(os.path.dirname(__file__), "config.json")) as f:
    config = json.load(f)
    
if __name__ == "__main__":
    EmailProcessor = EmailProcessor(config["APC_USER"], get_password(config["APC_USER"]))