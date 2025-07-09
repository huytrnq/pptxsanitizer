"""Text processing utilities."""

import re
from typing import List, Tuple, Dict, Any
import logging


class TextProcessor:
    """Utility class for text processing and normalization."""

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def normalize_text_for_matching(self, text: str) -> str:
        """Normalize text for better matching."""
        if not text:
            return ""

        # Remove extra whitespace
        text = re.sub(r"\s+", " ", text.strip())
        # Handle special characters
        text = text.replace("\u00a0", " ")  # Non-breaking space
        text = text.replace("\u2013", "-")  # En dash
        text = text.replace("\u2014", "-")  # Em dash
        text = text.replace("\u201c", '"')  # Left double quote
        text = text.replace("\u201d", '"')  # Right double quote
        text = text.replace("\u2018", "'")  # Left single quote
        text = text.replace("\u2019", "'")  # Right single quote

        return text

    def is_flexible_match(self, text: str, search_term: str) -> bool:
        """Check if search term matches with flexible whitespace/formatting."""
        # Create flexible pattern
        pattern = self.create_flexible_pattern(search_term)
        return bool(re.search(pattern, text, re.IGNORECASE))

    def create_flexible_pattern(self, search_term: str) -> str:
        """Create a flexible regex pattern for matching."""
        # Escape special regex characters
        escaped = re.escape(search_term)
        # Replace escaped spaces with flexible whitespace pattern
        flexible = escaped.replace(r"\ ", r"\s*")
        return flexible

    def apply_fuzzy_replacements(
        self, text: str, sorted_replacements: List[Tuple[str, str]]
    ) -> Dict[str, Any]:
        """Apply fuzzy/flexible text replacements."""
        new_text = text
        applied_replacements = []

        self.logger.info(f"Attempting fuzzy matching for: '{text}'")

        for original, replacement in sorted_replacements:
            if not original:
                continue

            # Normalize both texts
            normalized_text = self.normalize_text_for_matching(new_text)
            normalized_original = self.normalize_text_for_matching(original)

            # Try different matching strategies
            if normalized_original in normalized_text:
                # Simple normalized match
                new_text = new_text.replace(original, replacement)
                applied_replacements.append((original, replacement))
                self.logger.info(
                    f"  Fuzzy match (normalized): '{original}' -> '{replacement}'"
                )

            elif self.is_flexible_match(normalized_text, normalized_original):
                # Flexible regex match
                pattern = self.create_flexible_pattern(normalized_original)
                new_text = re.sub(
                    pattern, replacement, new_text, flags=re.IGNORECASE
                )
                applied_replacements.append((original, replacement))
                self.logger.info(
                    f"  Fuzzy match (flexible): '{original}' -> '{replacement}'"
                )

        return {"new_text": new_text, "replacements": applied_replacements}
