import os
import logging
from datetime import datetime

# Define the base log directory in the user's AppData/Local/ade bms/log directory
LOG_DIR = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "log")
os.makedirs(LOG_DIR, exist_ok=True)

# Create or reuse a log file with the current date in the name
current_date = datetime.now().strftime("%Y-%m-%d")
LOG_FILE_PATH = os.path.join(LOG_DIR, f"log_{current_date}.txt")

# Configure logging
logger = logging.getLogger("ade_bms_logger")
logger.setLevel(logging.DEBUG)

# File handler to write logs to a daily log file
file_handler = logging.FileHandler(LOG_FILE_PATH, mode='a')  # 'a' mode to append if file exists
file_handler.setLevel(logging.DEBUG)

# Log formatting including filename and line number
formatter = logging.Formatter(
    '%(asctime)s - %(levelname)s - %(message)s [%(filename)s:%(lineno)d]'
)
file_handler.setFormatter(formatter)

# Adding file handler to the logger
logger.addHandler(file_handler)

# Optional: Log to console as well
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

# Log the start of the application
logger.info("Application started.")
