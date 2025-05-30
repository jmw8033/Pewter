from EmailProcessor import EmailProcessor
import config
    
if __name__ == "__main__":
    EmailProcessor = EmailProcessor(config.APC_USER, config.APC_PASS)