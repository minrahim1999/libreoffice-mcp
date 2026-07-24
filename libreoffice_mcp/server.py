"""LibreOffice MCP Server — stdio transport for agentic code tools.

Exposes tools for creating and manipulating presentations, documents, and spreadsheets
via the Model Context Protocol (MCP). Works with Claude Code, Codex, OpenCode, and
any MCP-compatible agent.
"""
from __future__ import annotations

import json
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    ListToolsResult,
    TextContent,
    Tool,
)

from libreoffice_mcp.presentation import (
    create_presentation,
    add_slide,
    set_slide_title,
    add_bullet_points,
    add_table,
    add_image,
    export_pdf,
    get_presentation_info,
)
from libreoffice_mcp.document import (
    create_document,
    add_heading,
    add_paragraph,
    add_document_table,
)
from libreoffice_mcp.spreadsheet import (
    create_spreadsheet,
    add_sheet,
    set_cell_value,
    add_chart,
)
from libreoffice_mcp.utils import resolve_path

# ── server ──────────────────────────────────────────────────────────────────
server = Server("libreoffice-mcp")

# ── tool definitions ────────────────────────────────────────────────────────
TOOLS: list[Tool] = [
    # Presentations ---------------------------------------------------------
    Tool(
        name="presentation_create",
        description="Create a new .pptx presentation file.",
        inputSchema={
            "type": "object",
            "properties": {
                "output_path": {
                    "type": "string",
                    "description": "Absolute or relative path where the .pptx will be saved. Defaults to /tmp/presentation.pptx",
                },
                "width": {
                    "type": "number",
                    "description": "Slide width in inches. Default 13.333 (16:9 widescreen).",
                },
                "height": {
                    "type": "number",
                    "description": "Slide height in inches. Default 7.5 (16:9 widescreen).",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="presentation_add_slide",
        description="Add a slide to an existing .pptx. Returns the new slide index.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the existing .pptx file.",
                },
                "layout": {
                    "type": "string",
                    "enum": ["title", "title_and_content", "blank", "title_only"],
                    "description": "Slide layout type. Default: title_and_content",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="presentation_set_title",
        description="Set the title text of a specific slide.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "slide_index": {"type": "integer", "description": "0-based slide index."},
                "title": {"type": "string"},
                "font_size": {"type": "integer", "description": "Optional font size in points."},
                "bold": {"type": "boolean"},
            },
            "required": ["file_path", "slide_index", "title"],
        },
    ),
    Tool(
        name="presentation_add_bullets",
        description="Add bullet-point content to a slide's content placeholder.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "slide_index": {"type": "integer"},
                "bullets": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of bullet point strings.",
                },
                "font_size": {"type": "integer"},
            },
            "required": ["file_path", "slide_index", "bullets"],
        },
    ),
    Tool(
        name="presentation_add_table",
        description="Add a data table to a slide.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "slide_index": {"type": "integer"},
                "headers": {"type": "array", "items": {"type": "string"}},
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                },
                "left": {"type": "number", "description": "Position from left in inches. Default 1."},
                "top": {"type": "number", "description": "Position from top in inches. Default 2."},
                "width": {"type": "number", "description": "Table width in inches. Default 10."},
                "height": {"type": "number", "description": "Table height in inches. Default 3."},
            },
            "required": ["file_path", "slide_index", "headers", "rows"],
        },
    ),
    Tool(
        name="presentation_add_image",
        description="Add an image to a slide.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "slide_index": {"type": "integer"},
                "image_path": {"type": "string", "description": "Path to the image file."},
                "left": {"type": "number", "description": "Position in inches. Default 1."},
                "top": {"type": "number", "description": "Position in inches. Default 1."},
                "width": {"type": "number", "description": "Width in inches. Optional."},
                "height": {"type": "number", "description": "Height in inches. Optional."},
            },
            "required": ["file_path", "slide_index", "image_path"],
        },
    ),
    Tool(
        name="presentation_export_pdf",
        description="Export a .pptx presentation to PDF using LibreOffice headless mode.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "output_dir": {"type": "string", "description": "Optional output directory. Defaults to same dir as input."},
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="presentation_info",
        description="Get metadata about a .pptx file: slide count, dimensions, slide titles.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
            },
            "required": ["file_path"],
        },
    ),
    # Documents --------------------------------------------------------------
    Tool(
        name="document_create",
        description="Create a new .docx document.",
        inputSchema={
            "type": "object",
            "properties": {
                "output_path": {
                    "type": "string",
                    "description": "Path to save the .docx. Defaults to /tmp/document.docx",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="document_add_heading",
        description="Add a heading to a .docx document.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "text": {"type": "string"},
                "level": {"type": "integer", "description": "Heading level 0-9. Default 1.", "minimum": 0, "maximum": 9},
            },
            "required": ["file_path", "text"],
        },
    ),
    Tool(
        name="document_add_paragraph",
        description="Add a paragraph to a .docx document.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "text": {"type": "string"},
                "bold": {"type": "boolean"},
                "italic": {"type": "boolean"},
                "font_size": {"type": "integer"},
            },
            "required": ["file_path", "text"],
        },
    ),
    Tool(
        name="document_add_table",
        description="Add a table to a .docx document.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "headers": {"type": "array", "items": {"type": "string"}},
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                },
            },
            "required": ["file_path", "headers", "rows"],
        },
    ),
    # Spreadsheets -----------------------------------------------------------
    Tool(
        name="spreadsheet_create",
        description="Create a new .xlsx spreadsheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "output_path": {
                    "type": "string",
                    "description": "Path to save the .xlsx. Defaults to /tmp/spreadsheet.xlsx",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="spreadsheet_add_sheet",
        description="Add a new worksheet to an .xlsx file.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "title": {"type": "string", "description": "Sheet name."},
            },
            "required": ["file_path", "title"],
        },
    ),
    Tool(
        name="spreadsheet_set_cell",
        description="Set a cell value in a specific sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "sheet_title": {"type": "string"},
                "cell": {"type": "string", "description": "Cell reference, e.g. A1, B2."},
                "value": {"type": "string"},
            },
            "required": ["file_path", "sheet_title", "cell", "value"],
        },
    ),
    Tool(
        name="spreadsheet_add_chart",
        description="Add a chart to a spreadsheet sheet (requires data range).",
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "sheet_title": {"type": "string"},
                "chart_type": {
                    "type": "string",
                    "enum": ["bar", "column", "line", "pie"],
                    "description": "Chart type. Default: column",
                },
                "data_range": {"type": "string", "description": "Excel-style data range, e.g. A1:D5."},
                "title": {"type": "string", "description": "Chart title."},
                "position": {"type": "string", "description": "Cell position for chart top-left, e.g. F2."},
            },
            "required": ["file_path", "sheet_title", "data_range", "title"],
        },
    ),
]

# ── handlers ─────────────────────────────────────────────────────────────
@server.list_tools()
async def list_tools() -> ListToolsResult:
    return ListToolsResult(tools=TOOLS)

@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any] | None = None) -> list[TextContent]:
    args = arguments or {}
    try:
        result = await _dispatch(name, args)
        return [TextContent(type="text", text=result)]
    except Exception as exc:
        return [TextContent(type="text", text=f"ERROR: {type(exc).__name__}: {exc}")]


async def _dispatch(name: str, args: dict[str, Any]) -> str:
    """Route tool calls to the appropriate handler."""
    # ── presentations ──────────────────────────────────────────────────────
    if name == "presentation_create":
        path = resolve_path(args.get("output_path", "/tmp/presentation.pptx"))
        width = args.get("width", 13.333)
        height = args.get("height", 7.5)
        create_presentation(path, width=width, height=height)
        return json.dumps({"created": str(path), "width": width, "height": height})

    if name == "presentation_add_slide":
        path = resolve_path(args["file_path"])
        layout = args.get("layout", "title_and_content")
        idx = add_slide(path, layout)
        return json.dumps({"slide_index": idx, "layout": layout})

    if name == "presentation_set_title":
        path = resolve_path(args["file_path"])
        idx = args["slide_index"]
        title = args["title"]
        font_size = args.get("font_size")
        bold = args.get("bold")
        set_slide_title(path, idx, title, font_size=font_size, bold=bold)
        return json.dumps({"slide_index": idx, "title": title})

    if name == "presentation_add_bullets":
        path = resolve_path(args["file_path"])
        idx = args["slide_index"]
        bullets = args["bullets"]
        font_size = args.get("font_size")
        add_bullet_points(path, idx, bullets, font_size=font_size)
        return json.dumps({"slide_index": idx, "bullets_added": len(bullets)})

    if name == "presentation_add_table":
        path = resolve_path(args["file_path"])
        idx = args["slide_index"]
        headers = args["headers"]
        rows = args["rows"]
        left = args.get("left", 1.0)
        top = args.get("top", 2.0)
        width = args.get("width", 10.0)
        height = args.get("height", 3.0)
        add_table(path, idx, headers, rows, left=left, top=top, width=width, height=height)
        return json.dumps({"slide_index": idx, "columns": len(headers), "rows": len(rows)})

    if name == "presentation_add_image":
        path = resolve_path(args["file_path"])
        idx = args["slide_index"]
        image = resolve_path(args["image_path"])
        left = args.get("left", 1.0)
        top = args.get("top", 1.0)
        width = args.get("width")
        height = args.get("height")
        add_image(path, idx, image, left=left, top=top, width=width, height=height)
        return json.dumps({"slide_index": idx, "image": str(image)})

    if name == "presentation_export_pdf":
        path = resolve_path(args["file_path"])
        out = args.get("output_dir")
        pdf_path = export_pdf(path, output_dir=out)
        return json.dumps({"pdf": str(pdf_path)})

    if name == "presentation_info":
        path = resolve_path(args["file_path"])
        info = get_presentation_info(path)
        return json.dumps(info, indent=2)

    # ── documents ────────────────────────────────────────────────────────
    if name == "document_create":
        path = resolve_path(args.get("output_path", "/tmp/document.docx"))
        create_document(path)
        return json.dumps({"created": str(path)})

    if name == "document_add_heading":
        path = resolve_path(args["file_path"])
        level = args.get("level", 1)
        add_heading(path, args["text"], level=level)
        return json.dumps({"heading": args["text"], "level": level})

    if name == "document_add_paragraph":
        path = resolve_path(args["file_path"])
        add_paragraph(
            path,
            args["text"],
            bold=args.get("bold"),
            italic=args.get("italic"),
            font_size=args.get("font_size"),
        )
        return json.dumps({"paragraph": args["text"]})

    if name == "document_add_table":
        path = resolve_path(args["file_path"])
        add_document_table(path, args["headers"], args["rows"])
        return json.dumps({"table_added": True})

    # ── spreadsheets ───────────────────────────────────────────────────────
    if name == "spreadsheet_create":
        path = resolve_path(args.get("output_path", "/tmp/spreadsheet.xlsx"))
        create_spreadsheet(path)
        return json.dumps({"created": str(path)})

    if name == "spreadsheet_add_sheet":
        path = resolve_path(args["file_path"])
        add_sheet(path, args["title"])
        return json.dumps({"sheet": args["title"]})

    if name == "spreadsheet_set_cell":
        path = resolve_path(args["file_path"])
        set_cell_value(path, args["sheet_title"], args["cell"], args["value"])
        return json.dumps({"cell": args["cell"], "value": args["value"]})

    if name == "spreadsheet_add_chart":
        path = resolve_path(args["file_path"])
        sheet = args["sheet_title"]
        chart_type = args.get("chart_type", "column")
        data_range = args["data_range"]
        title = args["title"]
        position = args.get("position", "F2")
        add_chart(path, sheet, chart_type, data_range, title, position)
        return json.dumps({"chart": title, "type": chart_type})

    return json.dumps({"error": f"Unknown tool: {name}"})


# ── entry point ───────────────────────────────────────────────────────────
def main() -> None:
    import asyncio
    asyncio.run(_run())


async def _run() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    main()
