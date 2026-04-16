import logging

def setup_logging(verbose: bool) -> logging.Logger:
    """Configure root logger.

    Args:
        verbose: If True, use INFO level; otherwise WARNING.
    """
    level = logging.INFO if verbose else logging.WARNING
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    return logging.getLogger()
