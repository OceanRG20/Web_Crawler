import logging
import os

def setup_logger(name="crawler", log_file="output/crawler.log", level=logging.INFO):
    """Set up a logger that writes both to console and file."""
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Avoid duplicate handlers
    if not logger.handlers:
        # File handler
        fh = logging.FileHandler(log_file, mode="a", encoding="utf-8")
        fh.setLevel(level)

        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(level)

        # Formatter
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)

        logger.addHandler(fh)
        logger.addHandler(ch)

    return logger
