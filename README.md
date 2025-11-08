# Document Format MCP Server

MCP server for reading and writing Microsoft Office and Google Workspace document formats.

## Features

- ✅ Read PowerPoint (.pptx) files
- ✅ Read Word (.docx) files
- ✅ Read Excel (.xlsx) files
- ✅ Read Google Workspace files (Sheets, Docs, Slides)
- 🚧 Write PowerPoint (.pptx) files (Coming soon)
- 🚧 Write Word (.docx) files (Coming soon)
- 🚧 Write Excel (.xlsx) files (Coming soon)
- 🚧 Write Google Workspace files (Coming soon)

## Quick Start

### 5-Minute Test

```bash
# 1. Setup environment (2 min)
uv venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux
uv pip install -r requirements.txt

# 2. Generate sample files (1 min)
python create_sample_files.py

# 3. Run tests (2 min)
python test_readers.py
```

詳細は [QUICKSTART.md](QUICKSTART.md) を参照してください。

## Installation

### Using uvx (recommended)

```bash
uvx document-format-mcp-server
```

### Using pip

```bash
pip install document-format-mcp-server
```

## Configuration

### Kiro MCP Configuration

Add to `.kiro/settings/mcp.json`:

```json
{
  "mcpServers": {
    "document-format": {
      "command": "uvx",
      "args": ["document-format-mcp-server"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "~/.config/kiro-mcp/google-credentials.json",
        "MCP_OUTPUT_DIR": "~/Documents/kiro-output"
      },
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

### Configuration File

Create a `config.json` file (optional):

```json
{
  "google_credentials_path": "~/.config/kiro-mcp/google-credentials.json",
  "output_directory": "~/Documents/kiro-output",
  "max_file_size_mb": 100,
  "max_sheets": 100,
  "max_slides": 500,
  "api_timeout_seconds": 60,
  "enable_google_workspace": true
}
```

### Environment Variables

- `GOOGLE_APPLICATION_CREDENTIALS`: Path to Google API credentials file
- `MCP_OUTPUT_DIR`: Output directory for generated files
- `MCP_LOG_LEVEL`: Log level (DEBUG, INFO, WARNING, ERROR)
- `MCP_MAX_FILE_SIZE_MB`: Maximum file size in MB
- `MCP_API_TIMEOUT`: API timeout in seconds

## Google API Setup

To use Google Workspace features:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the following APIs:
   - Google Sheets API
   - Google Docs API
   - Google Slides API
4. Create OAuth 2.0 credentials
5. Download the credentials JSON file
6. Save it to the path specified in `google_credentials_path`

## Available Tools

### Read Tools

- `read_powerpoint`: Read PowerPoint (.pptx) files
- `read_word`: Read Word (.docx) files
- `read_excel`: Read Excel (.xlsx) files
- `read_google_spreadsheet`: Read Google Spreadsheets
- `read_google_document`: Read Google Documents
- `read_google_slides`: Read Google Slides

### Write Tools

- `write_powerpoint`: Create PowerPoint (.pptx) files
- `write_word`: Create Word (.docx) files
- `write_excel`: Create Excel (.xlsx) files
- `write_google_spreadsheet`: Create Google Spreadsheets
- `write_google_document`: Create Google Documents
- `write_google_slides`: Create Google Slides

## Development

### Local Development Setup

詳細なセットアップ手順は [SETUP.md](SETUP.md) を参照してください。

#### Using uv (Recommended)

```bash
# Create virtual environment
uv venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# Install dependencies
uv pip install -r requirements.txt
```

#### Using pip

```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

# Install dependencies
pip install -r requirements.txt
```

### Testing Reader Functions

#### Generate Sample Files

```bash
python create_sample_files.py
```

This creates:
- `test_files/sample.pptx` - PowerPoint file with 3 slides
- `test_files/sample.docx` - Word file with headings, paragraphs, and tables
- `test_files/sample.xlsx` - Excel file with 3 sheets

#### Run Tests

```bash
# Test local files (PowerPoint, Word, Excel)
python test_readers.py

# Test Google Workspace files
python test_readers.py --google
```

### Project Structure

```
document-format-mcp-server/
├── src/
│   └── document_format_mcp_server/
│       ├── server.py              # MCP server entry point
│       ├── readers/               # Document readers
│       │   ├── powerpoint_reader.py
│       │   ├── word_reader.py
│       │   ├── excel_reader.py
│       │   └── google_reader.py
│       ├── writers/               # Document writers (coming soon)
│       ├── tools/                 # MCP tool definitions (coming soon)
│       └── utils/                 # Utilities
│           ├── config.py
│           └── errors.py
├── tests/                         # Unit tests
├── test_files/                    # Test sample files
├── test_readers.py                # Reader function test script
├── create_sample_files.py         # Sample file generator
├── QUICKSTART.md                  # 5-minute quick start guide
├── SETUP.md                       # Detailed setup guide
├── requirements.txt               # Python dependencies
└── pyproject.toml                 # Project configuration
```

### Running Unit Tests

```bash
pytest
```

### Code Formatting

```bash
black src/ tests/
```

## Documentation

- [QUICKSTART.md](QUICKSTART.md) - 5分で動作確認できるクイックガイド
- [SETUP.md](SETUP.md) - 詳細なセットアップとトラブルシューティング
- [Design Document](.kiro/specs/document-format-mcp-server/design.md) - アーキテクチャと設計
- [Requirements](.kiro/specs/document-format-mcp-server/requirements.md) - 要件定義
- [Tasks](.kiro/specs/document-format-mcp-server/tasks.md) - 実装タスクリスト

## License

MIT
