#!/usr/bin/env python3
"""
PowerPoint Sanitizer - Main Entry Point
=======================================

Complete sanitization system for PowerPoint presentations.
"""

import sys
import logging
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.core.sanitizer import PowerPointSanitizer
from config import Config


def setup_logging():
    """Set up logging configuration."""
    logging.basicConfig(
        level=getattr(logging, Config.DEFAULT_LOG_LEVEL),
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )


def main():
    """Main sanitization workflow."""
    setup_logging()
    logger = logging.getLogger(__name__)

    # Configuration
    INPUT_FILE = str(Config.DEFAULT_INPUT_FILE)
    OPENAI_API_KEY = Config.get_openai_api_key()
    IMAGES_DIR = str(Config.IMAGES_DIR)
    PROMPTS_DIR = str(Config.PROMPTS_DIR)

    if not OPENAI_API_KEY:
        logger.error(
            "OpenAI API key not found. Please set OPENAI_API_KEY environment variable."
        )
        sys.exit(1)

    try:
        # Initialize sanitizer
        sanitizer = PowerPointSanitizer(
            openai_api_key=OPENAI_API_KEY, images_dir=IMAGES_DIR, prompts_dir=PROMPTS_DIR
        )

        # Generate output filename
        output_file = Config.get_output_filename(INPUT_FILE)

        # Run sanitization
        report = sanitizer.sanitize_presentation(INPUT_FILE, output_file)

        # Print summary
        sanitizer.print_summary(report)

        print(f"\nSanitization completed successfully!")
        print(f"Sanitized file: {report.sanitized_file}")
        print(f"Report file: {Path(report.sanitized_file).with_suffix('.json')}")

    except Exception as e:
        logger.error(f"Sanitization failed: {e}")
        raise


if __name__ == "__main__":
    main()
