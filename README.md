# PowerPoint Sanitizer

An AI-powered tool for automatically detecting and sanitizing sensitive information in PowerPoint presentations. This tool helps organizations safely share presentation materials by identifying and removing confidential data such as names, contacts, client-specific terms, and proprietary information.

## 🚀 Features

- **AI-Enhanced Detection**: Uses OpenAI's language models to intelligently identify sensitive content
- **Sanitization**: Removes multiple types of sensitive information
- **Detailed Reporting**: Generates comprehensive reports of all changes made

## 📋 Requirements

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
   # Windows PowerShell
   $env:OPENAI_API_KEY = "your-api-key-here"
   
   # Linux/Mac
   export OPENAI_API_KEY="your-api-key-here"
   ```

## 🚀 Quick Start

1. **Place your PowerPoint file** in the `data/` directory (default: `Take-home.pptx`)

2. **Run the sanitizer:**

   ```bash
   python main.py
   ```

3. **Find your sanitized file** in the `data/` directory with `_sanitized` suffix

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

The tool uses environment variables and configuration files:

- **OPENAI_API_KEY**: Your OpenAI API key (required)
- **Input file**: Default is `data/Take-home.pptx`
- **Prompts**: Customize AI behavior by editing files in `config/prompts/`

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

## 🔧 Development

### Adding Custom Prompts

1. Edit files in `config/prompts/`
2. Modify `system_prompt.txt` for AI behavior
3. Update `user_prompt.txt` for detection instructions

