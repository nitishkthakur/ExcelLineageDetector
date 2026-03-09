import logging
import sys


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("%(levelname)s [%(name)s] %(message)s"))
        logger.addHandler(handler)
    return logger


def set_log_level(verbose: bool):
    level = logging.DEBUG if verbose else logging.INFO
    logging.getLogger("lineage").setLevel(level)
