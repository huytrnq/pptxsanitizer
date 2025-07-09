"""Models package for PowerPoint sanitizer."""

from .slide_data import SlideData
from .detection import Detection
from .sanitization_report import SanitizationReport

__all__ = ["SlideData", "Detection", "SanitizationReport"]
