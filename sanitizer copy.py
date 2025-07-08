from typing import List, Dict, Any
from dataclasses import dataclass, field

from pptx_parser import PPTXParser
from openai_analyzer import OpenAIAnalyzer

if __name__ == "__main__":
    # Example usage
    pptx_parser = PPTXParser()
    openai_analyzer = OpenAIAnalyzer(
        api_key="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"
    )

    # Load a PowerPoint file
    slides_data = pptx_parser.parse_presentation("data/Take-home.pptx")
    for i, slide in enumerate(slides_data):
        print(f"Slide {i + 1}: {slide.title}")
        print(f"Text content: {slide.text_content}")
        print(f"Images count: {slide.images_count}")
        print(f"Charts count: {slide.charts_count}")
        print(f"Tables count: {slide.tables_count}")

        # Analyze the slide text with OpenAI
        slide_image_path = f"data/pngs/slide_{i + 1:02d}.png"
        sanitized_content = openai_analyzer.analyze_slide(
            slide.text_content, slide_image_path
        )
        print(f"Sanitized content: {sanitized_content}\n")
        # break
