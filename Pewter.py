from EmailProcessor import EmailProcessor
import threading
import config
import time
    
if __name__ == "__main__":
    # ACP
    acp_email_processor = threading.Thread(target=EmailProcessor, args=(config.ACP_USER, config.ACP_PASS))
    acp_email_processor.start()
    time.sleep(0.1)

    # APC
    apc_email_processor = threading.Thread(target=EmailProcessor, args=(config.APC_USER, config.APC_PASS))
    apc_email_processor.start()