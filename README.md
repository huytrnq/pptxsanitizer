# PowerPoint Sanitization System

A modular system for detecting and sanitizing sensitive information in PowerPoint presentations using AI-enhanced detection capabilities.

## Quick Start

1. **Extract the provided zip file** to your working directory

2. **Install dependencies**:
   ```bash
   uv sync
   ```

3. **Run sanitization**:
   ```bash
   # Basic mode (pattern-only)
   python main.py test_presentation.pptx sanitized_output.pptx
   
   # With AI (requires OpenAI API key)
   python main.py test_presentation.pptx sanitized_output.pptx --api-key "your-key-here"
   ```

## File Structure

```
pptx_sanitizer/
├── main.py                     # Main orchestrator
├── pptx_handler.py            # PowerPoint file operations
├── ai_processor.py            # AI content analysis
├── pyproject.toml             # Project configuration and dependencies    
└── README.md                  # This documentation
```

## Usage

### Basic Usage

```bash
python main.py input.pptx output_sanitized.pptx
```

### With OpenAI API Key

```bash
python main.py input.pptx output_sanitized.pptx --api-key "your-api-key"
```

### Pattern-Only Mode (No AI)

```bash
python main.py input.pptx output_sanitized.pptx --no-ai
```

### Advanced Options

```bash
python main.py input.pptx output_sanitized.pptx \
    --api-key "your-api-key" \
    --confidence 0.8 \
    --report sanitization_report.json
```

### Command Line Arguments

- `input_file`: Path to input PowerPoint file (.pptx)
- `output_file`: Path for sanitized output file (.pptx)
- `--api-key`: OpenAI API key for AI-enhanced detection
- `--no-ai`: Disable AI detection (pattern-only mode)
- `--confidence`: Confidence threshold for replacements (0.0-1.0, default: 0.7)
- `--report`: Path for detailed sanitization report (JSON)

## Configuration

### OpenAI API Key

Set your API key in one of these ways:

1. **Environment variable**:
   ```bash
   export OPENAI_API_KEY="your-key-here"
   python main.py input.pptx output.pptx
   ```

2. **Command line**:
   ```bash
   python main.py input.pptx output.pptx --api-key "your-key-here"
   ```

3. **Pattern-only mode** (no API key needed):
   ```bash
   python main.py input.pptx output.pptx --no-ai
   ```

## Detection Categories

### Pattern-Based Detection (Always Active)
- **Email addresses**: `user@domain.com` → `[EMAIL]`
- **Phone numbers**: `+1-555-123-4567` → `[PHONE]`
- **Company names**: `Acme Corporation` → `[CLIENT]`
- **Personal names**: `John Smith` → `[NAME]`
- **Financial data**: `$50.2 million` → `[AMOUNT]`

### AI-Enhanced Detection (When API Key Provided)
- **Client identifiers**: Company-specific terms and branding
- **Internal terminology**: Acronyms and internal processes
- **Geographic locations**: Specific addresses and regional identifiers
- **Product names**: Proprietary products and technologies
- **Confidential information**: Classified or sensitive content
