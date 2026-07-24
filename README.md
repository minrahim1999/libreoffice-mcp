# LibreOffice MCP Server

A Model Context Protocol (MCP) server that gives agentic coding tools programmatic control over **presentations** (pptx), **documents** (docx), and **spreadsheets** (xlsx) — plus PDF export via LibreOffice headless mode.

Works with any MCP-compatible agent: **Claude Code**, **Codex**, **OpenCode**, **Pi Agent**, **Cursor**, **Windsurf**, etc.

---

## Tools Overview

### Presentations (`presentation_*`)

| Tool | Description |
|------|-------------|
| `presentation_create` | Create a new `.pptx` with configurable dimensions |
| `presentation_add_slide` | Add a slide (title / title_and_content / blank / title_only) |
| `presentation_set_title` | Set a slide's title text |
| `presentation_add_bullets` | Add bullet points to a slide |
| `presentation_add_table` | Add a styled data table to a slide |
| `presentation_add_image` | Add an image to a slide |
| `presentation_export_pdf` | Convert `.pptx` → `.pdf` via LibreOffice |
| `presentation_info` | Get slide count, dimensions, titles |

### Documents (`document_*`)

| Tool | Description |
|------|-------------|
| `document_create` | Create a new `.docx` |
| `document_add_heading` | Add a heading at any level |
| `document_add_paragraph` | Add formatted paragraph text |
| `document_add_table` | Add a table with headers + rows |

### Spreadsheets (`spreadsheet_*`)

| Tool | Description |
|------|-------------|
| `spreadsheet_create` | Create a new `.xlsx` |
| `spreadsheet_add_sheet` | Add a worksheet |
| `spreadsheet_set_cell` | Write a value to a cell |
| `spreadsheet_add_chart` | Add a bar / column / line / pie chart |

---

## Prerequisites

- **Python 3.10+**
- **LibreOffice** (for PDF export — `soffice` on PATH)
- **pip** or **uv**

Install LibreOffice:

```bash
# macOS
brew install --cask libreoffice

# Ubuntu/Debian
sudo apt install libreoffice

# Fedora
sudo dnf install libreoffice
```

---

## Installation

### Option A: pip (editable dev install)

```bash
cd ~/Documents/01-Projects/libreoffice-mcp
pip install -e .
```

### Option B: uv (recommended)

```bash
cd ~/Documents/01-Projects/libreoffice-mcp
uv pip install -e .
```

This installs the `libreoffice-mcp` CLI entrypoint.

---

## Agent Configuration (Tested Locally)

### Claude Code

```bash
# ~/.claude/mcp.json
cat > ~/.claude/mcp.json <<'EOF'
{
  "mcpServers": {
    "libreoffice": {
      "command": "/Users/muhaimin/.hermes/hermes-agent/venv/bin/libreoffice-mcp"
    }
  }
}
EOF
```

Restart Claude Code with `Ctrl+Shift+P` → "Restart Claude Code Server".

### OpenCode (`~/.config/opencode/opencode.jsonc`)

Add under `mcpServers`:

```jsonc
    "libreoffice": {
      "type": "local",
      "command": [
        "/Users/muhaimin/.hermes/hermes-agent/venv/bin/libreoffice-mcp"
      ],
      "enabled": true
    }
```

Restart OpenCode.

### Hermes Agent (`~/.hermes/config.yaml`)

Add under `mcp_servers`:

```yaml
  libreoffice:
    command: /Users/muhaimin/.hermes/hermes-agent/venv/bin/libreoffice-mcp
    enabled: true
```

Restart Hermes.

---

## Quickstart Example

Create a presentation with 3 slides programmatically:

```python
# Agent calls these via MCP tools

1. presentation_create {"output_path": "/tmp/deck.pptx"}
2. presentation_set_title  {"file_path": "/tmp/deck.pptx", "slide_index": 0, "title": "Quarterly Review"}
3. presentation_add_bullets {"file_path": "/tmp/deck.pptx", "slide_index": 0, "bullets": ["Revenue +24%", "8 releases shipped", "1.2M users"]}
4. presentation_add_slide {"file_path": "/tmp/deck.pptx", "layout": "title_and_content"}
5. presentation_set_title {"file_path": "/tmp/deck.pptx", "slide_index": 1, "title": "Roadmap"}
6. presentation_add_bullets {"file_path": "/tmp/deck.pptx", "slide_index": 1, "bullets": ["Q3: Foundation", "Q4: Acceleration", "Q1: Scale"]}
7. presentation_export_pdf {"file_path": "/tmp/deck.pptx"}
```

---

## Project Structure

```
libreoffice-mcp/
├── pyproject.toml
├── README.md
└── libreoffice_mcp/
    ├── __init__.py
    ├── server.py          # MCP stdio server + tool dispatch
    ├── presentation.py    # pptx operations (python-pptx)
    ├── document.py        # docx operations (python-docx)
    ├── spreadsheet.py     # xlsx operations (openpyxl)
    └── utils.py           # Path helpers
```

---

## Troubleshooting

**`soffice` not found when exporting PDF**

Make sure LibreOffice is installed and `soffice` is on your PATH:

```bash
which soffice
soffice --version
```

On macOS, Homebrew sometimes links it as `/opt/homebrew/bin/soffice`. Ensure your shell PATH includes this.

**MCP tools not appearing**

- Check the agent's MCP server logs for connection errors.
- Verify `libreoffice-mcp` runs standalone: `echo '{}' | libreoffice-mcp`
- Ensure Python dependencies are installed: `pip install python-pptx openpyxl python-docx mcp`

---

## License

MIT

## Template Tools (Human-Grade Output)

For professional, styled output without manual formatting:

| Tool | Description | Best For |
|------|-------------|----------|
| `template_executive_presentation` | 5-slide corporate deck with hero, metrics, financials, roadmap | Board decks, quarterly reviews |
| `template_proposal_document` | Multi-page doc with title page, sections, budget table | Client proposals, SOWs |
| `template_financial_spreadsheet` | Styled Excel with formulas, margins, embedded chart | Monthly/quarterly reports |

These templates produce **navy/orange color scheme**, **16:9 widescreen**, and **professional typography** — no manual tweaking needed.
