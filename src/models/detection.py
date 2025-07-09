"""Detection model."""

from dataclasses import dataclass
from pydantic import BaseModel
from typing import List


@dataclass
class Detection:
    """Detection object for compatibility."""

    original: str
    replacement: str
    category: str = ""
    reason: str = ""


class OpenAIDetection(BaseModel):
    """A single sensitive content detection with enhanced details."""

    original: str
    replacement: str
    category: str
    reason: str
    sensitivity_level: str = "MEDIUM"  # HIGH/MEDIUM/LOW


class DetectionResponse(BaseModel):
    """Response containing multiple detections."""

    detections: List[OpenAIDetection]
