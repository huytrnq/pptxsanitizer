"""Configuration module."""

import os
from pathlib import Path


class Config:
    """Configuration settings for the PowerPoint sanitizer."""

    # File paths
    DATA_DIR = Path("data")
    IMAGES_DIR = DATA_DIR / "pngs"
    PROMPTS_DIR = Path("config") / "prompts"
    
    # Default files
    DEFAULT_INPUT_FILE = DATA_DIR / "Take-home.pptx"
    DEFAULT_OUTPUT_SUFFIX = "_sanitized"
    
    # OpenAI settings
    DEFAULT_MODEL = "gpt-4.1-mini-2025-04-14"
    DEFAULT_TEMPERATURE = 0.4
    DEFAULT_MAX_TOKENS = 4000
    
    # Logging
    DEFAULT_LOG_LEVEL = "INFO"
    
    @classmethod
    def get_openai_api_key(cls) -> str:
        """Get OpenAI API key from environment variable."""
        return os.getenv("OPENAI_API_KEY", "")
    
    @classmethod
    def get_output_filename(cls, input_file: str) -> str:
        """Generate output filename from input filename."""
        input_path = Path(input_file)
        return str(
            input_path.parent /
            f"{input_path.stem}{cls.DEFAULT_OUTPUT_SUFFIX}{input_path.suffix}"
        )
