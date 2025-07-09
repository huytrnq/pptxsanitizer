"""
PowerPoint Sanitizer
===================

A complete system for finding and removing sensitive information from PowerPoint files.

The sanitizer works in 4 steps:
1. Extract text and images from PowerPoint slides
2. Use AI to detect sensitive information 
3. Replace sensitive text with safe alternatives
4. Generate a report of what was changed

Supports detecting:
- Personal information (names, emails, phones)
- Financial data (account numbers, SSNs)
- Business confidential information
- Healthcare data
- Custom sensitive patterns
"""

import os
import json
import logging
from typing import List, Dict, Any
from pathlib import Path

from .pptx_processor import PPTXProcessor
from .openai_analyzer import OpenAIAnalyzer
from ..models.slide_data import SlideData
from ..models.detection import Detection
from ..models.sanitization_report import SanitizationReport


class PowerPointSanitizer:
    """
    Sanitizes PowerPoint files by finding and replacing sensitive information.
    
    Uses AI to analyze slides and detect sensitive content like personal data,
    financial information, and confidential business data. Creates a clean
    version of the presentation with sensitive text replaced.
    """

    def __init__(
        self,
        openai_api_key: str,
        images_dir: str = "data/pngs",
        prompts_dir: str = "config/prompts",
        model: str = "gpt-4.1-mini-2025-04-14"
    ):
        """
        Initialize the sanitizer.
        
        Args:
            openai_api_key: Your OpenAI API key
            images_dir: Folder with slide images (default: "data/pngs")
            prompts_dir: Folder with AI prompts (default: "config/prompts")
            model: OpenAI model to use (default: "gpt-4.1-mini-2025-04-14")
        """
        self.pptx_processor = PPTXProcessor()
        self.analyzer = OpenAIAnalyzer(api_key=openai_api_key, 
                                    prompts_dir=prompts_dir,
                                    model=model)
        self.images_dir = Path(images_dir)
        self.logger = logging.getLogger(__name__)


    def sanitize_presentation(
        self, input_file: str, output_file: str = None
    ) -> SanitizationReport:
        """
        Sanitize a PowerPoint file by removing sensitive information.

        Args:
            input_file: Path to the PowerPoint file to sanitize
            output_file: Where to save the clean file (optional)

        Returns:
            SanitizationReport: Summary of what was found and changed
        """
        # Set default output file name
        if not output_file:
            input_path = Path(input_file)
            output_file = str(
                input_path.parent / f"{input_path.stem}_sanitized{input_path.suffix}"
            )

        self.logger.info(f"Starting sanitization of {input_file}")

        # 1. Shape Identification & Extraction
        slides_data = self.pptx_processor.parse_presentation(input_file)
        self.logger.info(f"Extracted data from {len(slides_data)} slides")

        # 2. AI-Enhanced Sensitive Data Detection
        all_detections = {}
        for slide in slides_data:
            slide_image_path = self.images_dir / f"slide_{slide.slide_number:02d}.png"

            if slide_image_path.exists():
                detections = self._analyze_slide(slide, slide_image_path)
                all_detections[slide.slide_number] = detections
                detection_count = (
                    len(detections.detections)
                    if hasattr(detections, "detections")
                    else len(detections)
                )
                self.logger.info(
                    f"Slide {slide.slide_number}: {detection_count} detections"
                )
            else:
                self.logger.warning(
                    f"Image not found for slide {slide.slide_number}: {slide_image_path}"
                )
                all_detections[slide.slide_number] = []

        # 3. Content Replacement
        processed_detections = self._convert_detections_for_replacement(all_detections)

        # Apply all replacements to file
        replacement_result = self.pptx_processor.apply_replacements_to_file(
            input_file, output_file, processed_detections
        )

        # Handle both old (bool) and new (dict) return types
        if isinstance(replacement_result, bool):
            # Old return type - just boolean success
            replacement_success = replacement_result
            total_replacements = 0  # Can't get count from old version
            if not replacement_success:
                self.logger.error("Failed to apply replacements")
        else:
            # New return type - detailed results dictionary
            replacement_success = replacement_result.get("success", False)
            total_replacements = replacement_result.get("total_replacements", 0)
            if not replacement_success:
                self.logger.error(
                    f"Failed to apply replacements: {replacement_result.get('error', 'Unknown error')}"
                )

        self.logger.info(
            f"Replacement process completed. Total replacements: {total_replacements}"
        )

        # 4. Generate Report with actual replacement count
        report = self._generate_report(
            input_file, output_file, slides_data, all_detections, total_replacements
        )
        self._save_report(report, output_file)

        self.logger.info(f"Sanitization completed. Output: {output_file}")
        return report

    def _analyze_slide(self, slide_data: SlideData, image_path: Path):
        """
        Analyze one slide for sensitive content using AI.
        
        Args:
            slide_data: Text and metadata from the slide
            image_path: Path to the slide image file
                
        Returns:
            List of detected sensitive information or empty list
        """
        self.logger.info(f"Analyzing slide {slide_data.slide_number}")
        self.logger.debug(f"  Text content: {slide_data.text_content}")
        self.logger.debug(f"  Image path: {image_path}")
        self.logger.debug(f"  Image exists: {image_path.exists()}")

        if not slide_data.text_content:
            self.logger.warning(
                f"Slide {slide_data.slide_number}: No text content to analyze"
            )
            return []

        try:
            detections = self.analyzer.analyze_slide(
                slide_text=slide_data.text_content, image_path=str(image_path)
            )

            detection_count = (
                len(detections.detections)
                if hasattr(detections, "detections")
                else len(detections)
            )
            self.logger.info(
                f"Slide {slide_data.slide_number}: Analysis returned {detection_count} detections"
            )

            return detections
        except Exception as e:
            self.logger.error(f"Error analyzing slide {slide_data.slide_number}: {e}")
            return []

    def _convert_detections_for_replacement(
        self, all_detections: Dict[int, Any]
    ) -> Dict[int, List[Detection]]:
        """
        Convert AI detection results to format needed for text replacement.
        
        Args:
            all_detections: Dictionary of detections from AI analysis
                
        Returns:
            Dictionary of standardized Detection objects for replacement
        """
        processed_detections = {}

        for slide_number, detections in all_detections.items():
            if hasattr(detections, "detections"):
                detection_list = detections.detections
            else:
                detection_list = detections if isinstance(detections, list) else []

            # Convert to Detection format
            pptx_detections = []
            for detection in detection_list:
                pptx_detection = Detection(
                    original=getattr(
                        detection, "original", getattr(detection, "text", "")
                    ),
                    replacement=getattr(detection, "replacement", ""),
                    category=getattr(detection, "category", "unknown"),
                    reason=getattr(detection, "reason", ""),
                )
                pptx_detections.append(pptx_detection)

            processed_detections[slide_number] = pptx_detections

        return processed_detections

    def _generate_report(
        self,
        input_file: str,
        output_file: str,
        slides_data: List[SlideData],
        all_detections: Dict[int, Any],
        total_replacements: int = 0,
    ) -> SanitizationReport:
        """
        Create a summary report of the sanitization process.
        
        Args:
            input_file: Original PowerPoint file path
            output_file: Sanitized PowerPoint file path
            slides_data: Data from all slides
            all_detections: All sensitive content found
            total_replacements: Number of text replacements made
                
        Returns:
            SanitizationReport: Summary with statistics and details
        """

        # Count total detections and categorize
        total_detections = 0
        categories_summary = {}
        detections_by_slide = {}

        for slide_number, detections in all_detections.items():
            if hasattr(detections, "detections"):
                detection_list = detections.detections
            else:
                detection_list = detections if isinstance(detections, list) else []

            detections_by_slide[slide_number] = detection_list
            total_detections += len(detection_list)

            # Count categories
            for detection in detection_list:
                category = getattr(detection, "category", "unknown")
                categories_summary[category] = categories_summary.get(category, 0) + 1

        return SanitizationReport(
            original_file=input_file,
            sanitized_file=output_file,
            total_slides=len(slides_data),
            total_detections=total_detections,
            total_replacements=total_replacements,
            detections_by_slide=detections_by_slide,
            categories_summary=categories_summary,
        )

    def _save_report(self, report: SanitizationReport, output_file: str):
        """
        Save the report as a JSON file.
        
        Args:
            report: The sanitization report to save
            output_file: PowerPoint file path (JSON will have same name)
        """
        report_file = Path(output_file).with_suffix(".json")

        # Convert report to serializable format
        report_data = {
            "original_file": report.original_file,
            "sanitized_file": report.sanitized_file,
            "total_slides": report.total_slides,
            "total_detections": report.total_detections,
            "total_replacements": getattr(report, "total_replacements", 0),
            "categories_summary": report.categories_summary,
            "detections_by_slide": {
                str(slide_num): [
                    {
                        "original": getattr(d, "original", getattr(d, "text", "")),
                        "replacement": getattr(d, "replacement", ""),
                        "category": getattr(d, "category", "unknown"),
                        "reason": getattr(d, "reason", ""),
                        "sensitivity_level": getattr(d, "sensitivity_level", "MEDIUM"),
                    }
                    for d in detections
                ]
                for slide_num, detections in report.detections_by_slide.items()
            },
        }

        with open(report_file, "w", encoding="utf-8") as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)

        self.logger.info(f"Saved sanitization report to {report_file}")

    def print_summary(self, report: SanitizationReport):
        """
        Print a summary of sanitization results to the console.
        
        Args:
            report: The sanitization report to display
        """
        print(f"\n=== SANITIZATION SUMMARY ===")
        print(f"Original file: {report.original_file}")
        print(f"Sanitized file: {report.sanitized_file}")
        print(f"Total slides: {report.total_slides}")
        print(f"Total detections: {report.total_detections}")
        print(f"Total replacements: {getattr(report, 'total_replacements', 'N/A')}")

        print(f"\nDetections by category:")
        for category, count in report.categories_summary.items():
            print(f"  {category}: {count}")

        print(f"\nDetections by slide:")
        for slide_num, detections in report.detections_by_slide.items():
            if detections:
                print(f"  Slide {slide_num}: {len(detections)} detections")
