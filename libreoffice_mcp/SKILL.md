# LibreOffice MCP — Agent Instructions

When using the `libreoffice-mcp` tools, follow these conventions for professional, human-grade output.

## General Principles

- **Prefer templates over raw tools.** Use `template_executive_presentation`, `template_proposal_document`, and `template_financial_spreadsheet` for polished output. Only use granular tools (`presentation_add_slide`, `document_add_paragraph`, etc.) when the user needs custom control.
- **Always provide complete data.** Templates expect structured JSON. Partial data produces partial slides.
- **Use absolute paths** for `output_path` so the user can find the file.

## Presentation Templates

### `template_executive_presentation`

Best for: quarterly reviews, board decks, strategic overviews.

**Required:**
- `title`, `presenter`, `date`

**Recommended:**
- `metrics`: 3 items max, each with `label`, `value`, `description`
- `financials`: table with `headers` and `rows`
- `roadmap`: up to 3 phases with `phase`, `quarter`, `details`

**Output:** 5 professionally designed slides (hero, metrics, financials, roadmap, thank-you) with navy/orange color scheme and 16:9 widescreen format.

### `template_proposal_document`

Best for: client proposals, SOWs, project briefs.

**Required:**
- `title`, `client`, `author`, `date`

**Recommended:**
- `summary`: 2-3 sentence executive summary
- `sections`: array of `{heading, paragraphs}` for each section
- `budget_table`: `{headers, rows}` for cost breakdown

**Output:** Multi-page Word document with title page, formatted headings, and styled tables.

### `template_financial_spreadsheet`

Best for: monthly/quarterly financial reports with charts.

**Required:**
- `title`, `months`, `revenue`, `expenses`, `profit`

**Output:** Styled Excel file with colored headers, calculated margins, and embedded column chart.

## Style Guidelines

- **Titles:** Use sentence case, not Title Case ("Q3 executive summary" not "Q3 Executive Summary")
- **Metrics:** Use plain numbers in `value` ("24%" not "+24% YoY"), keep context in `description`
- **Tables:** Keep column headers short (2-3 words). Right-align currency columns.
- **Roadmap:** Use verb-noun format for phases ("Build foundation" not "Foundation")

## Export

Always offer to `presentation_export_pdf` after creating a pptx so the user has a shareable PDF version.
