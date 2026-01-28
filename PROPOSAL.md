# Proposal: Draw.io to Editable PowerPoint Converter

## 1. Executive Summary
Develop a Python-based application to convert Draw.io (`.drawio`, `.xml`) diagrams into editable Microsoft PowerPoint (`.pptx`) presentations. The key differentiator is preserving the editability of text, shapes, and specifically connectors (lines/arrows) within PowerPoint.

## 2. Technical Approach

### 2.1 Core Technologies
- **Language:** Python 3.12+ (compatible with `uv` manager)
- **Primary Libraries:**
    - `python-pptx`: For generating `.pptx` files, managing slides, shapes, and connectors.
    - `drawio2pptx` (Open Source): Evaluate as a base library. It already implements mxGraph to PresentationML translation.
    - `defusedxml`: For safe XML parsing of Draw.io files.

### 2.2 Conversion Strategy
The conversion process will involve three main stages:

1.  **Parsing (mxGraph XML):**
    - Extract the `<mxGraphModel>` from the input file.
    - Parse `mxGeometry` to determine position (x, y) and dimensions (width, height).
    - Parse `style` attributes to determine shape type (rectangle, ellipse, etc.), fill color, stroke color, and font styling.
    - Identify connections: `source` and `target` attributes in `<mxCell>` elements interact with `python-pptx` connector objects.

2.  **Mapping & Generation:**
    - **Shapes:** Map standard Draw.io shapes (General, Flowchart) to `python-pptx` AutoShapes (`MSO_SHAPE`).
    - **Text:** Convert HTML-formatted labels from Draw.io into PowerPoint TextFrames, preserving font size, bold/italic, and alignment where possible.
    - **Connectors:** Crucially, instead of drawing static lines, we will use `slide.shapes.add_connector()`. We will programmatically link the connector's start/end to the specific connection points of the source/target shapes using `connector.begin_connect()` and `connector.end_connect()`. This ensures that moving a shape in PowerPoint drags the connected line with it.

3.  **Layout & Scaling:**
    - Map Draw.io pixel coordinates to PowerPoint EMUs (English Metric Units). 1 pixel ≈ 9525 EMUs.
    - Handle multi-page diagrams by creating multiple slides.

## 3. Implementation Plan

### Phase 1: Feasibility & Prototype (Days 1-2)
- [ ] Set up project environment with `uv`.
- [ ] Create a "Hello World" script that reads a minimal Draw.io XML (two rectangles + one arrow) and generates a PPTX.
- [ ] Verify connector "stickiness" (editability) in the output.

### Phase 2: Core Converter (Days 3-5)
- [ ] Implement parsing for common shapes: Rectangles, Ellipses, Diamonds, Rounded Rectangles.
- [ ] Implement text styling: Font size, color, bold/italic.
- [ ] Handle standard arrows and line styles (solid, dashed).

### Phase 3: Web Application (Day 6)
- [ ] Develop a lightweight web interface (similar to Pdf2Markdown).
- [ ] Drag-and-drop file upload.
- [ ] Instant download of converted `.pptx`.

## 4. Known Limitations & Mitigations
- **Custom Shapes:** Complex SVG-based shapes in Draw.io may not have a direct equivalent in PowerPoint.
    - *Mitigation:* Render these as high-resolution PNG images if an editable shape cannot be approximated.
- **Rich Text:** Complex HTML formatting in Draw.io labels might lose some fidelity (e.g., nested lists).
    - *Mitigation:* Simplify to plain text with basic formatting (bold/italic) or parse basic HTML tags to PPTX text runs.

## 5. Directory Structure
```
~/WebApps/Tools/Drawio2PPTX/
├── main.py
├── converter/
│   ├── parser.py
│   ├── pptx_generator.py
│   └── styles.py
├── webapp/
│   └── app.py
├── tests/
├── pyproject.toml
└── README.md
```
