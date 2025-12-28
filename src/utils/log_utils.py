"""
Centralized logging utility for AlSawifeFactory application.
Only logs errors to keep the output clean and focused.
"""

import logging
import os
from datetime import datetime

# Create logs directory in Documents
LOGS_DIR = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
os.makedirs(LOGS_DIR, exist_ok=True)

# Log file path
LOG_FILE = os.path.join(LOGS_DIR, "app_errors.log")

# Configure the logger
logger = logging.getLogger("AlSawifeFactory")
logger.setLevel(logging.ERROR)

# Prevent duplicate handlers
if not logger.handlers:
    # File handler - logs to file
    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    file_handler.setLevel(logging.ERROR)
    
    # Console handler - logs to console
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.ERROR)
    
    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s | %(levelname)s | %(module)s:%(lineno)d | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)


def log_error(message: str, exc_info: bool = False):
    """
    Log an error message.
    
    Args:
        message: The error message to log
        exc_info: If True, include exception traceback
    """
    logger.error(message, exc_info=exc_info)


def log_exception(message: str):
    """
    Log an error with full exception traceback.
    
    Args:
        message: The error message to log
    """
    logger.exception(message)


def get_log_file_path() -> str:
    """Return the path to the log file."""
    return LOG_FILE
