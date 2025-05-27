from EmailProcessor import EmailProcessor
from Rectangulator import RectangulatorHandler
import threading
import config
import time
    
if __name__ == "__main__":
    print("""
██████╗ ███████╗██╗    ██╗████████╗███████╗██████╗ 
██╔══██╗██╔════╝██║    ██║╚══██╔══╝██╔════╝██╔══██╗
██████╔╝█████╗  ██║ █╗ ██║   ██║   █████╗  ██████╔╝
██╔═══╝ ██╔══╝  ██║███╗██║   ██║   ██╔══╝  ██╔══██╗
██║     ███████╗╚███╔███╔╝   ██║   ███████╗██║  ██║
╚═╝     ╚══════╝ ╚══╝╚══╝    ╚═╝   ╚══════╝╚═╝  ╚═╝                                       
""")
    EmailProcessor = EmailProcessor(config.APC_USER, config.APC_PASS)