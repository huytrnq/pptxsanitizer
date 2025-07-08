import json
import logging
import base64
import os
from typing import List, Dict, Any
from pydantic import BaseModel
from openai import OpenAI

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class Detection(BaseModel):
    """A single sensitive content detection."""
    original: str
    category: str
    replacement: str
    reason: str 


class DetectionResponse(BaseModel):
    """Response containing multiple detections."""
    detections: List[Detection]

class OpenAIAnalyzer:
    """Analyzes text content for sensitive information using OpenAI."""

    def __init__(self, api_key: str, model="gpt-4o-mini", prompts_dir="prompts"):
        """Initialize with OpenAI API key and prompt directory."""
        if not api_key:
            raise ValueError("OpenAI API key is required")

        self.client = OpenAI(api_key=api_key)
        self.logger = logging.getLogger(__name__)
        self.model = model
        self.prompts_dir = prompts_dir

        # Load prompts from files
        self.system_prompt = self._load_prompt("system_prompt.txt")
        self.user_prompt = self._load_prompt("user_prompt.txt")

    def _load_prompt(self, filename: str) -> str:
        """Load a prompt from a file."""
        try:
            prompt_path = os.path.join(self.prompts_dir, filename)
            with open(prompt_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            self.logger.error("Error loading prompt %s: %s", filename, e)
            return ""

    def _encode_image(self, image_path: str) -> str:
        """Encode image to base64 string."""
        try:
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode("utf-8")
        except Exception as e:
            self.logger.error("Error encoding image %s: %s", image_path, e)
            raise

    def _prepare_system_prompt(self, slide_text: List[str]) -> str:
        """Prepare the system prompt with the extracted text."""
        # Format the text content as a proper list representation
        if isinstance(slide_text, list):
            formatted_text = str(slide_text)
        else:
            formatted_text = slide_text
        return self.system_prompt.format(extracted_text_list=formatted_text)

    def analyze_slide(self, slide_text: List[str], image_path: str) -> str:
        """Analyze a single slide's text content for sensitive information."""
        try:
            self.logger.info(
                "Analyzing slide with text length: %d characters", len(slide_text)
            )

            # Encode the image
            base64_image = self._encode_image(image_path)

            # Prepare the system prompt with extracted text
            system_prompt = self._prepare_system_prompt(slide_text)

            # Create the messages for OpenAI API
            messages = [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": self.user_prompt},
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
            return response_content

        except Exception as e:
            self.logger.error("Error analyzing slide: %s", e)
            raise


    def get_sanitization_summary(self, detections: List[Detection]) -> Dict[str, Any]:
        """Get a summary of sanitization results."""
        categories = {}
        for detection in detections:
            if detection.category not in categories:
                categories[detection.category] = 0
            categories[detection.category] += 1

        return {
            "total_detections": len(detections),
            "categories": categories,
            "detections": [
                {
                    "original": d.text,
                    "replacement": d.replacement,
                    "category": d.category,
                }
                for d in detections
            ],
        }
