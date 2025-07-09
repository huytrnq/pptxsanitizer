"""Slide data model."""

from typing import List
from dataclasses import dataclass, field


@dataclass
class SlideData:
    """Data from a single slide."""

    slide_number: int
    title: str = ""
    text_content: List[str] = field(default_factory=list)
    images_count: int = 0
    charts_count: int = 0
    tables_count: int = 0
