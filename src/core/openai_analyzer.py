"""OpenAI analyzer for sensitive content detection."""

import json
import logging
import base64
import os
from typing import List, Dict, Any

from openai import OpenAI

from ..models.detection import OpenAIDetection, DetectionResponse


class OpenAIAnalyzer:
    """Analyzes text content for sensitive information using OpenAI with improved prompts."""

    def __init__(
        self, api_key: str, model="gpt-4.1-mini-2025-04-14", prompts_dir="config/prompts"
    ):
        """Initialize with OpenAI API key and prompt directory."""
        if not api_key:
            raise ValueError("OpenAI API key is required")

        self.client = OpenAI(api_key=api_key)
        self.logger = logging.getLogger(__name__)
        self.model = model
        self.prompts_dir = prompts_dir

        # Load prompts from files or use improved defaults
        self.system_prompt = self._load_prompt("system_prompt.txt")
        self.user_prompt = self._load_prompt("user_prompt.txt")

    def _load_prompt(self, filename: str) -> str:
        """Load a prompt from a file."""
        try:
            prompt_path = os.path.join(self.prompts_dir, filename)
            with open(prompt_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            self.logger.warning("Could not load prompt %s: %s", filename, e)
            return ""

    def _encode_image(self, image_path: str) -> str:
        """Encode image to base64 string."""
        try:
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode("utf-8")
        except Exception as e:
            self.logger.error("Error encoding image %s: %s", image_path, e)
            raise

    def _prepare_user_prompt(self, slide_text: List[str]) -> str:
        """Prepare the system prompt with the extracted text."""
        # Format the text content as a proper list representation
        if isinstance(slide_text, list):
            formatted_text = str(slide_text)
        else:
            formatted_text = slide_text
        return self.user_prompt.format(extracted_text_list=formatted_text)

    def analyze_slide(
        self, slide_text: List[str], image_path: str
    ) -> DetectionResponse:
        """Analyze a single slide's text content for sensitive information."""
        try:
            self.logger.info(
                "Analyzing slide with %d text elements", len(slide_text)
            )

            # Encode the image
            base64_image = self._encode_image(image_path)

            # Prepare the system prompt with extracted text
            user_prompt = self._prepare_user_prompt(slide_text)

            # Create the messages for OpenAI API
            messages = [
                {"role": "system", "content": self.system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": user_prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            },
                        },
                    ],
                },
            ]

            # Make the API call
            response = self.client.chat.completions.parse(
                model=self.model,
                messages=messages,
                max_tokens=4000,
                temperature=0.1,  # Low temperature for consistent sanitization
                response_format=DetectionResponse,
            )

            # Parse the response
            response_content = response.choices[0].message.parsed

            # Log summary of detections
            if response_content and response_content.detections:
                high_risk = sum(
                    1
                    for d in response_content.detections
                    if d.sensitivity_level == "HIGH"
                )
                medium_risk = sum(
                    1
                    for d in response_content.detections
                    if d.sensitivity_level == "MEDIUM"
                )
                low_risk = sum(
                    1
                    for d in response_content.detections
                    if d.sensitivity_level == "LOW"
                )

                self.logger.info(
                    f"Found {len(response_content.detections)} detections: "
                    f"{high_risk} HIGH risk, {medium_risk} MEDIUM risk, "
                    f"{low_risk} LOW risk"
                )

            return response_content

        except Exception as e:
            self.logger.error("Error analyzing slide: %s", e)
            raise

    def get_sanitization_summary(
        self, detections: List[OpenAIDetection]
    ) -> Dict[str, Any]:
        """Get a summary of sanitization results."""
        categories = {}
        sensitivity_levels = {}

        for detection in detections:
            # Count categories
            if detection.category not in categories:
                categories[detection.category] = 0
            categories[detection.category] += 1

            # Count sensitivity levels
            if detection.sensitivity_level not in sensitivity_levels:
                sensitivity_levels[detection.sensitivity_level] = 0
            sensitivity_levels[detection.sensitivity_level] += 1

        return {
            "total_detections": len(detections),
            "categories": categories,
            "sensitivity_levels": sensitivity_levels,
            "detections": [
                {
                    "original": d.original,
                    "replacement": d.replacement,
                    "category": d.category,
                    "reason": d.reason,
                    "sensitivity_level": d.sensitivity_level,
                }
                for d in detections
            ],
        }
