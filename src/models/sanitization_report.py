"""Sanitization report model."""

from typing import List, Dict
from dataclasses import dataclass, field
from .detection import Detection


@dataclass
class SanitizationReport:
    """Report of sanitization results."""

    original_file: str
    sanitized_file: str
    total_slides: int
    total_detections: int
    total_replacements: int = 0
    detections_by_slide: Dict[int, List[Detection]] = field(
        default_factory=dict
    )
    categories_summary: Dict[str, int] = field(default_factory=dict)
