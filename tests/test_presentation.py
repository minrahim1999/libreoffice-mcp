"""Tests for presentation module."""
import json
import tempfile
from pathlib import Path

from libreoffice_mcp.presentation import (
    create_presentation,
    add_slide,
    set_slide_title,
    add_bullet_points,
    get_presentation_info,
)


def test_create_and_info():
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        path = Path(f.name)
    try:
        create_presentation(path, width=10, height=7.5)
        info = get_presentation_info(path)
        assert info["slide_count"] == 1
        assert abs(info["width_inches"] - 10) < 0.01
        assert abs(info["height_inches"] - 7.5) < 0.01
    finally:
        path.unlink(missing_ok=True)


def test_add_slide_and_bullets():
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        path = Path(f.name)
    try:
        create_presentation(path)
        idx = add_slide(path, "title_and_content")
        assert idx == 1
        set_slide_title(path, 1, "Test Slide")
        add_bullet_points(path, 1, ["Point A", "Point B"])
        info = get_presentation_info(path)
        assert info["slide_count"] == 2
        assert info["slides"][1]["title"] == "Test Slide"
    finally:
        path.unlink(missing_ok=True)
