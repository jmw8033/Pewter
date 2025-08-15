from EmailProcessor import EmailProcessor
import json, os

with open(os.path.join(os.path.dirname(__file__), "config.json")) as f:
    config = json.load(f)
    
if __name__ == "__main__":
    EmailProcessor = EmailProcessor(config["APC_USER"], config["APC_PASS"])