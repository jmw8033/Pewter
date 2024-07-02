from EmailProcessor import EmailProcessor
from Rectangulator import RectangulatorHandler
import threading
import config
import time
    
if __name__ == "__main__":
    # Rectangulator Handler
    rectangulator_handler = RectangulatorHandler()

    # ACP
    acp_email_processor = threading.Thread(target=EmailProcessor, args=(config.ACP_USER, config.ACP_PASS, rectangulator_handler))
    acp_email_processor.start()
    time.sleep(0.1)

    # APC
    apc_email_processor = threading.Thread(target=EmailProcessor, args=(config.APC_USER, config.APC_PASS, rectangulator_handler))
    apc_email_processor.start()

    print("""
██████╗ ███████╗██╗    ██╗████████╗███████╗██████╗ 
██╔══██╗██╔════╝██║    ██║╚══██╔══╝██╔════╝██╔══██╗
██████╔╝█████╗  ██║ █╗ ██║   ██║   █████╗  ██████╔╝
██╔═══╝ ██╔══╝  ██║███╗██║   ██║   ██╔══╝  ██╔══██╗
██║     ███████╗╚███╔███╔╝   ██║   ███████╗██║  ██║
╚═╝     ╚══════╝ ╚══╝╚══╝    ╚═╝   ╚══════╝╚═╝  ╚═╝                                       
""")