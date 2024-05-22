from EmailProcessor import EmailProcessor
import threading
import config
    
if __name__ == "__main__":
    # ACP
    acp_email_processor = threading.Thread(target=EmailProcessor, args=(config.ACP_USER, config.ACP_PASS))
    acp_email_processor.start()

    # APC
    apc_email_processor = threading.Thread(target=EmailProcessor, args=(config.APC_USER, config.APC_PASS))
    apc_email_processor.start()