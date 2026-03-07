# vibe-office

> Chat-based AI agent for Excel and Word file operations — powered by Claude, OpenRouter, or Ollama.

[日本語版 README はこちら](README.ja.md)

## Features

- **Excel** — read/write cells, ranges, and formulas; manage sheets; apply formatting
- **Word** — read documents, insert/delete/replace paragraphs, add headings, tables, and images
- **Multi-provider** — works with Anthropic Claude, Google Gemini, OpenRouter (cloud), and Ollama (local/offline)
- **Smart tool selection** — automatically passes only Excel or Word tools to the model based on context, reducing token usage for local models
- **Animated spinner** — shows the active tool name while the agent is thinking

## Requirements

- Python 3.11+
- [uv](https://docs.astral.sh/uv/)
- API key for your chosen provider

## Installation

```bash
git clone <repo-url>
cd vibe-office
uv sync
```

## Configuration

### Using a .env file (recommended)

Copy `.env.example` to `.env` and fill in your API keys. The `.env` file is listed in `.gitignore` and **will never be committed to Git**.

```bash
cp .env.example .env
```

Example `.env` contents:

```
ANTHROPIC_API_KEY=sk-ant-...
GEMINI_API_KEY=AIza...
OPENROUTER_API_KEY=sk-or-...
```

Load it before running:

```bash
# bash / zsh
source .env && uv run python main.py

# Or use direnv for automatic loading
```

### Using environment variables directly

```bash
# Anthropic (default)
export ANTHROPIC_API_KEY="sk-ant-..."

# Google Gemini
# Get your key at https://aistudio.google.com/app/apikey
export GEMINI_API_KEY="AIza..."

# OpenRouter
export OPENROUTER_API_KEY="sk-or-..."

# Ollama — no key required (runs locally)
```

> **Warning**: Never hardcode API keys in source code. Always use environment variables or a `.env` file.

## Usage

```bash
# Anthropic Claude (default)
uv run python main.py

# OpenRouter
uv run python main.py --provider openrouter

# OpenRouter with a specific model
uv run python main.py --provider openrouter --model google/gemini-flash-1.5

# Google Gemini
uv run python main.py --provider gemini

# Google Gemini with a specific model
uv run python main.py --provider gemini --model gemini-2.5-pro

# Ollama (local)
uv run python main.py --provider ollama

# Ollama with a specific model or custom endpoint
uv run python main.py --provider ollama --model llama3.2
uv run python main.py --provider ollama --base-url http://192.168.1.10:11434/v1
```

You can also use environment variables instead of flags:

```bash
export EXCEL_AGENT_PROVIDER=ollama
export EXCEL_AGENT_MODEL=qwen2.5:7b
uv run python main.py
```

## Chat examples

```
You> Open sales.xlsx and show me the contents of Sheet1
You> Write "Revenue" in A1 and 250000 in B1
You> Put a SUM formula in C1 for the range B1:B10
You> Make the header row bold with a blue background
You> Create a new sheet called "Summary"

You> Open report.docx and show me the document structure
You> Replace "Q3 results" with "Q3 & Q4 results"
You> Insert a paragraph after index 3 saying "Additional notes follow."
You> Insert logo.png at the top of the document, 12 cm wide
You> Add a table with quarterly sales data
```

## Built-in commands

| Command | Description |
|---------|-------------|
| `/reset` | Clear conversation history |
| `/ls [path]` | List files (Excel and Word files highlighted) |
| `/cd <path>` | Change working directory |
| `/cwd` | Show current working directory |
| `/quit` | Exit |

## Supported providers & default models

| Provider | Flag | Default model | API key env var |
|----------|------|---------------|----------------|
| Anthropic | `--provider anthropic` | `claude-sonnet-4-6` | `ANTHROPIC_API_KEY` |
| Google Gemini | `--provider gemini` | `gemini-2.0-flash` | `GEMINI_API_KEY` |
| OpenRouter | `--provider openrouter` | `anthropic/claude-3.5-sonnet` | `OPENROUTER_API_KEY` |
| Ollama | `--provider ollama` | `qwen2.5:7b` | not required |

## Excel tools (13)

| Tool | Description |
|------|-------------|
| `open_excel` | Open or create a workbook |
| `list_sheets` | List all sheet names |
| `read_sheet` | Read sheet contents with optional range |
| `read_cell` | Read a single cell value |
| `write_cell` | Write a value to a cell |
| `write_range` | Write a 2D array starting from a cell |
| `apply_formula` | Set a formula in a cell |
| `create_sheet` | Add a new sheet |
| `delete_sheet` | Remove a sheet |
| `format_cell` | Apply bold, color, background, alignment |
| `set_column_width` | Set column width |
| `save_excel` | Save or save-as |
| `get_sheet_info` | Get row/column counts and used range |

## Word tools (13)

| Tool | Description |
|------|-------------|
| `open_word` | Open or create a document |
| `read_document` | Read full content with paragraph indices |
| `read_paragraph` | Read a single paragraph and its formatting |
| `append_paragraph` | Append a paragraph at the end |
| `insert_paragraph` | Insert before or after a paragraph index |
| `replace_text` | Find and replace text |
| `delete_paragraph` | Delete a paragraph by index |
| `format_paragraph` | Apply formatting to a paragraph |
| `insert_image` | Insert an image file |
| `add_table` | Insert a table from a 2D array |
| `add_heading` | Append a heading (level 1–9) |
| `save_word` | Save or save-as |
| `get_document_info` | Get heading list, paragraph and table counts |

## Project structure

```
vibe-office/
├── main.py          # Chat UI and CLI entry point
├── agent.py         # Claude API / OpenAI-compatible agent loop
├── excel_tools.py   # Excel operations (openpyxl)
├── word_tools.py    # Word operations (python-docx)
├── pyproject.toml   # Project config (uv)
├── .env.example     # API key template
└── .env             # Your API keys (git-ignored, create from .env.example)
```
