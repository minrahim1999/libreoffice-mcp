# Contributing

Thanks for your interest in improving libreoffice-mcp!

## Quick Start

```bash
git clone https://github.com/muhaimin/libreoffice-mcp.git
cd libreoffice-mcp
pip install -e ".[dev]"
pytest
```

## Adding a Tool

1. Implement the logic in the appropriate module (`presentation.py`, `document.py`, `spreadsheet.py`)
2. Add the `Tool` definition to `server.py` in the `TOOLS` list
3. Wire it in `_dispatch` in `server.py`
4. Add a test in `tests/`
5. Update `README.md` tool table

## Code Style

- Follow PEP 8
- Use type hints
- Keep functions focused and testable

## Pull Request Process

1. Fork and branch (`feature/your-change`)
2. Add tests
3. Ensure CI passes
4. Open PR with a clear description
