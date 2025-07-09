"""Logging utilities."""

import logging


def setup_logging(level=logging.INFO):
    """Set up logging configuration."""
    logging.basicConfig(level=level)
    return logging.getLogger(__name__)
