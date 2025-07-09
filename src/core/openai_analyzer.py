"""
OpenAI Analyzer for Sensitive Content Detection
===============================================

This module provides an analyzer that uses OpenAI's vision and language models
to detect sensitive information in PowerPoint slides. It processes both text
content and visual elements to identify potentially sensitive data that should
be sanitized or redacted.

The analyzer supports:
- Multi-modal analysis (text + images)
- Configurable sensitivity levels
- Custom prompts for domain-specific detection
- Structured response parsing with categorized detections
"""

import json
import logging
import base64
import os
from typing import List, Dict, Any

from openai import OpenAI

from ..models.detection import OpenAIDetection, DetectionResponse
from config import Config


class OpenAIAnalyzer:
    """
    Analyzes text content for sensitive information using OpenAI with improved prompts.
    
    This class leverages OpenAI's multimodal capabilities to analyze both textual content
    and visual elements of PowerPoint slides. It identifies sensitive information such as
    personal data, financial information, confidential business data, and other potentially
    sensitive content that may need to be sanitized.
    
    The analyzer uses structured prompts and response parsing to provide detailed detection
    results with categorization, sensitivity levels, and suggested replacements.
    
    Attributes:
        client: OpenAI API client instance
        logger: Logger for tracking analysis operations
        model: OpenAI model name to use for analysis
        prompts_dir: Directory containing custom prompt templates
        temperature: Sampling temperature for model responses
        max_tokens: Maximum tokens for model responses
        system_prompt: System prompt template for analysis
        user_prompt: User prompt template for specific requests
    """

    def __init__(
        self, 
        api_key: str, 
        model=None, 
        prompts_dir=None,
        temperature=None,
        max_tokens=None
    ):
        """
        Initialize the OpenAI analyzer with API credentials and configuration.
        
        Sets up the OpenAI client, loads custom prompts if available, and configures
        model parameters for sensitive content detection.
        
        Args:
            api_key (str): OpenAI API key for authentication
            model (str, optional): OpenAI model name. Defaults to Config.DEFAULT_MODEL
            prompts_dir (str, optional): Directory containing prompt templates. 
                Defaults to Config.PROMPTS_DIR
            temperature (float, optional): Sampling temperature (0.0-2.0). 
                Defaults to Config.DEFAULT_TEMPERATURE
            max_tokens (int, optional): Maximum response tokens. 
                Defaults to Config.DEFAULT_MAX_TOKENS
                
        Raises:
            ValueError: If API key is not provided
        """
        if not api_key:
            raise ValueError("OpenAI API key is required")

        self.client = OpenAI(api_key=api_key)
        self.logger = logging.getLogger(__name__)
        self.model = model or Config.DEFAULT_MODEL
        self.prompts_dir = prompts_dir or str(Config.PROMPTS_DIR)
        self.temperature = temperature or Config.DEFAULT_TEMPERATURE
        self.max_tokens = max_tokens or Config.DEFAULT_MAX_TOKENS

        # Load prompts from files or use improved defaults
        self.system_prompt = self._load_prompt("system_prompt.txt")
        self.user_prompt = self._load_prompt("user_prompt.txt")

    def _load_prompt(self, filename: str) -> str:
        """
        Load a prompt template from a file.
        
        Args:
            filename (str): Name of the prompt file to load
            
        Returns:
            str: Content of the prompt file, or empty string if loading fails
        """
        try:
            prompt_path = os.path.join(self.prompts_dir, filename)
            with open(prompt_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            self.logger.warning("Could not load prompt %s: %s", filename, e)
            return ""

    def _encode_image(self, image_path: str) -> str:
        """
        Encode image file to base64 string for API transmission.
        
        Reads the image file and converts it to a base64-encoded string suitable
        for sending to OpenAI's vision API endpoints.
        
        Args:
            image_path (str): Path to the image file to encode
            
        Returns:
            str: Base64-encoded string representation of the image
            
        Raises:
            Exception: If image file cannot be read or encoded
        """
        try:
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode("utf-8")
        except Exception as e:
            self.logger.error("Error encoding image %s: %s", image_path, e)
            raise

    def _prepare_user_prompt(self, slide_text: List[str]) -> str:
        """
        Prepare the user prompt with extracted slide text content.
        
        Formats the extracted text content into the user prompt template,
        ensuring proper string representation for API consumption.
        
        Args:
            slide_text (List[str]): List of text strings extracted from the slide
            
        Returns:
            str: Formatted user prompt with embedded text content
        """
        # Format the text content as a proper list representation
        if isinstance(slide_text, list):
            formatted_text = str(slide_text)
        else:
            formatted_text = slide_text
        return self.user_prompt.format(extracted_text_list=formatted_text)

    def analyze_slide(
        self, slide_text: List[str], image_path: str
    ) -> DetectionResponse:
        """
        Analyze a single slide's content for sensitive information.
        
        Performs multimodal analysis using both the extracted text content and
        the visual representation of the slide. Uses OpenAI's vision capabilities
        to identify sensitive information that may not be captured in text extraction.
        
        Args:
            slide_text (List[str]): List of text strings extracted from the slide
            image_path (str): Path to the slide image file for visual analysis
            
        Returns:
            DetectionResponse: Structured response containing detected sensitive
                information with categories, sensitivity levels, and replacements
                
        Raises:
            Exception: If API call fails or image cannot be processed
        """
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
                max_tokens=self.max_tokens,
                temperature=self.temperature,
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
        """
        Generate a comprehensive summary of sanitization results.
        
        Processes a list of detection results to create statistical summaries
        and organized views of the sensitive content found. Useful for reporting
        and understanding the scope of sanitization needed.
        
        Args:
            detections (List[OpenAIDetection]): List of detection objects from analysis
            
        Returns:
            Dict[str, Any]: Summary dictionary containing:
                - total_detections (int): Total number of detections
                - categories (Dict[str, int]): Count of detections by category
                - sensitivity_levels (Dict[str, int]): Count by sensitivity level
                - detections (List[Dict]): Detailed list of all detections with
                original text, replacement, category, reason, and sensitivity level
        """
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
