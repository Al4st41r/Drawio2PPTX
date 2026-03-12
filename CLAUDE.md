# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Drawio2PPTX converts Draw.io XML diagrams (`.drawio`, `.xml`) into fully editable PowerPoint presentations. Shapes, text, and connectors are native PowerPoint objects, not images.

## Commands

```bash
# Install dependencies
uv sync

# CLI conversion
uv run main.py input.drawio output.pptx

# Run web app (development)
uv run webapp/app.py

# Verify a generated PPTX
uv run tests/verify_output.py <pptx-file>
```

## Architecture

The conversion pipeline flows: **Draw.io XML → parsed data structures → PowerPoint PPTX**.

### `converter/` module

- **`__init__.py`** — Exports `convert(input_path, output_path)`, orchestrates parser then generator.
- **`parser.py`** (`DrawioParser`) — Extracts `<mxGraphModel>` from `<diagram>` elements. Separates vertices (shapes) from edges (connectors). Recursively resolves absolute positions for grouped shapes. Returns per-page `(vertices, edges)` tuples.
- **`engine.py`** (`PptxGenerator`) — The largest module. Creates slides using blank layout (index 6). Adds shapes mapped via `ppt_map`, applies styling, then routes connectors between shapes. Connector routing uses alignment heuristics (vertical/horizontal/diagonal) to select connection points and connector type (STRAIGHT vs ELBOW). Connector geometry is set via direct OXML xfrm manipulation rather than `python-pptx`'s `begin_connect`/`end_connect` (which produces incorrect bounding boxes that LibreOffice misrenders). The `stCxn`/`endCxn` connection references are intentionally omitted for the same reason. Edge labels are positioned using shape centres with support for Draw.io relative geometry (`x` as position along edge, `y` as perpendicular offset) and explicit pixel offsets from `<mxPoint as="offset">`.
- **`ppt_map.py`** — Static mappings from Draw.io style names to `MSO_SHAPE` enums, dash styles, and arrow types. Unsupported shapes fall back to rectangles.
- **`utils.py`** — Pixel-to-EMU conversion (`1px ≈ 9525 EMU`), hex-to-RGB, style string parsing, `HtmlTextParser` for extracting bold/italic/underline/colour/size from HTML labels, and OXML injection for arrowheads (not exposed by `python-pptx` high-level API).

### `webapp/` module

Flask app on port 5003 with a single POST endpoint accepting file uploads. Production deployment uses systemd (`drawio2pptx.service`) behind Nginx reverse proxy at `/drawio2pptx/`.

## Key Technical Details

- **Coordinate system:** Draw.io uses pixels; PowerPoint uses EMU. All geometry passes through `px_to_emu()`.
- **Arrowheads:** Set via direct OXML manipulation (`set_line_end()` in utils.py) because `python-pptx` lacks high-level arrowhead support. Draw.io's `startArrow` maps to OXML `headEnd` (source) and `endArrow` maps to OXML `tailEnd` (target).
- **Connector geometry:** Set via direct OXML xfrm to avoid `python-pptx` bounding-box issues and LibreOffice misrendering. Connection references (`stCxn`/`endCxn`) are omitted because LibreOffice recalculates geometry from them, overriding the explicit coordinates.
- **Connector routing:** The alignment detection and connection-point selection logic in `engine.py` (lines ~205–280) is the most complex and actively evolving part of the codebase.
- **Text frames:** All shapes use zero internal margins and `TEXT_TO_FIT_SHAPE` auto-sizing to maximise usable text area, particularly important for diamond shapes.
- **HTML text parsing:** `HtmlTextParser` handles `<b>`, `<i>`, `<font>`, `<span>` tags. Nested lists and some edge cases are unsupported.
- **XML security:** Uses `defusedxml` for all XML parsing.

## Dependencies

Python >=3.13, managed with `uv`. Key libraries: `python-pptx`, `lxml`, `defusedxml`, `flask`.
