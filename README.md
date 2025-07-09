# PowerPoint Sanitizer

An AI-powered tool for automatically detecting and sanitizing sensitive information in PowerPoint presentations. This tool helps organizations safely share presentation materials by identifying and removing confidential data such as names, contacts, client-specific terms, and proprietary information.

## 🚀 Features

- **AI-Enhanced Detection**: Uses OpenAI's language models to intelligently identify sensitive content
- **Sanitization**: Removes multiple types of sensitive information
- **Detailed Reporting**: Generates comprehensive reports of all changes made

## 📋 Requirements
- uv package and project manager
- Python 3.12 or higher
- OpenAI API key
- PowerPoint files (.pptx format)

### Dependencies

- **python-pptx**: PowerPoint file manipulation
- **openai**: AI-powered content analysis
- **typing-extensions**: Enhanced type annotations
- **requests**: HTTP requests handling


## 🛠️ Installation

1. **Install dependencies using uv:**

   ```bash
   uv sync
   ```

2. **Set up your OpenAI API key:**

   ```bash
   export OPENAI_API_KEY="your-api-key-here"
   ```

## 🚀 Quick Start

1. Convert your PowerPoint file to png images (Optional as the pngs are generated in the data/pngs directory):
   If you want to convert your PowerPoint slides to images, you can use `libreoffice` and `imagemagick`. Make sure you have them installed:
   ```bash
   libreoffice --headless --convert-to pdf data/Take-home.pptx --outdir data/
   sudo apt install imagemagick
   mkdir -p data/pngs
   convert -density 300 data/Take-home.pdf -scene 1 -quality 90 data/pngs/slide-%02d.png
   ```
2. **Place your PowerPoint file** in the `data/` directory (default: `Take-home.pptx`)

3. **Run the sanitizer:**

   ```bash
   python main.py
   ```

4. **Find your sanitized file** in the `data/` directory with `_sanitized` suffix

## 📁 Project Structure

```text
pptxsanitizer/
├── main.py                         # Main code for running the sanitizer
├── config/                         # Configuration files
│   ├── prompts/                    # AI prompt templates
│   │   ├── system_prompt.txt
│   │   └── user_prompt.txt
│   └── __init__.py
├── src/                            # Source code
│   ├── core/                       # Core functionality
│   │   ├── sanitizer.py            # Main sanitization logic
│   │   ├── pptx_processor.py       # PowerPoint file handling
│   │   └── openai_analyzer.py      # AI analysis
│   ├── models/                     # Data models
│   │   ├── detection.py            # Detection result data structures
│   │   ├── sanitization_report.py  # Report data structures
│   │   └── slide_data.py           # Slide data structures
│   └── utils/                      # Utility functions
│       ├── log.py                  # Logging utilities
│       └── text_processing.py      # Text processing helpers
├── data/                           # Input/output files
│   ├── pngs/                       # Slide images (if needed)
│   └── *.pptx                      # PowerPoint files
└── pyproject.toml                  # Project configuration
```

## ⚙️ Configuration

The tool uses a centralized configuration system managed by the `Config` class in `config/__init__.py`. This provides default settings and environment-based configuration for all components.

### Configuration Settings

The `Config` class includes the following default settings:

**File Paths:**

- `DATA_DIR`: `data/` - Directory for input/output files
- `IMAGES_DIR`: `data/pngs/` - Directory for slide images
- `PROMPTS_DIR`: `config/prompts/` - Directory for AI prompt templates
- `DEFAULT_INPUT_FILE`: `data/Take-home.pptx` - Default PowerPoint file to process

**OpenAI Settings:**

- `DEFAULT_MODEL`: `gpt-4.1-mini-2025-04-14` - Default AI model
- `DEFAULT_TEMPERATURE`: `0.1` - Controls AI response randomness (lower = more consistent)
- `DEFAULT_MAX_TOKENS`: `4000` - Maximum tokens per API request

**Output Settings:**

- `DEFAULT_OUTPUT_SUFFIX`: `_sanitized` - Suffix added to sanitized files

### Environment Variables

- **OPENAI_API_KEY**: Your OpenAI API key (required)

  ```bash
  export OPENAI_API_KEY="your-api-key-here"
  ```

### Customizing Configuration

You can customize the tool's behavior by:

1. **Modifying prompts**: Edit files in `config/prompts/`
   - `system_prompt.txt`: Defines AI assistant behavior
   - `user_prompt.txt`: Contains detection instructions

2. **Using the Config class**: Access default settings in your code

   ```python
   from config import Config
   
   # Get default model
   model = Config.DEFAULT_MODEL
   
   # Generate output filename
   output_file = Config.get_output_filename("input.pptx")
   
   # Get API key
   api_key = Config.get_openai_api_key()
   ```

## 📖 Usage Examples

```python
from src.core.sanitizer import PowerPointSanitizer

# Initialize sanitizer
sanitizer = PowerPointSanitizer(
   model="gpt-4.1-mini-2025-04-14",
   openai_api_key="your-api-key",
   images_dir="data/pngs",
   prompts_dir="config/prompts"
)

# Sanitize presentation
report = sanitizer.sanitize_presentation("input.pptx", "output_sanitized.pptx")

# Print summary
sanitizer.print_summary(report)
```

## 🛡️ Sanitization Guidelines

The tool follows comprehensive sanitization guidelines to remove:

1. **Names and Contacts**: Personal names, logos, emails, phone numbers
2. **Client-Specific Terms**: Acronyms, internal terminology, unique taxonomies
3. **Hidden Connections**: Subsidiaries, partners, vendors, suppliers
4. **Market Context**: Market position, competitive landscape, geographical identifiers
5. **Non-Public Insights**: Commercial strategies, cost structures, production data
6. **Visuals**: Client-specific charts, hidden identifiers, proprietary designs (Not Supported Yet)

## 📊 Output

The sanitizer generates:

- **Sanitized PowerPoint file**: Clean version with sensitive data removed
- **JSON report**: Detailed log of all detections and changes made
