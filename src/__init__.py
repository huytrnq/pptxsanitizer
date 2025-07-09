"""
PowerPoint Sanitizer Package
============================

A comprehensive system for sanitizing PowerPoint presentations by detecting
and replacing sensitive information using AI analysis.
"""

__version__ = "1.0.0"
__author__ = "PowerPoint Sanitizer Team"

# Make main components available at package level
from src.core import PPTXProcessor, OpenAIAnalyzer, PowerPointSanitizer
from src.models import SlideData, Detection, SanitizationReport

__all__ = [
    "PPTXProcessor",
    "OpenAIAnalyzer", 
    "PowerPointSanitizer",
    "SlideData",
    "Detection",
    "SanitizationReport"
]
