import logging
from pathlib import Path


def setup_logger(log_file: Path, level=logging.DEBUG):
    """Setup the root logger."""
    # Configure the root logger
    logger = logging.getLogger()
    logger.setLevel(level)

    # Avoid duplicate handlers if setup_logger is called multiple times
    if not logger.hasHandlers():
        # Define formatter
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] [%(filename)s:%(funcName)s]: %(message)s",
            datefmt="%d.%m.%Y %H:%M:%S"
        )

        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)

        # File handler
        file_handler = logging.FileHandler(log_file, mode="w")
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)

        # Add handlers to the logger
        logger.addHandler(console_handler)
        logger.addHandler(file_handler)
