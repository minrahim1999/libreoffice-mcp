"""Utility helpers for LibreOffice MCP."""
from __future__ import annotations

from pathlib import Path


def resolve_path(p: str) -> Path:
    """Resolve a path string — expand user and make absolute."""
    return Path(p).expanduser().resolve()
