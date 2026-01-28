# Planned Improvements (GitHub Issues)

## Issue 1: Expanded Shape Library Support
**Type:** Feature
**Priority:** High

**Description:**
Currently, the converter supports a core set of shapes (Rectangle, Ellipse, Diamond, etc.). Many standard Draw.io shapes fall back to default rectangles or aren't mapped correctly.

**Tasks:**
- [ ] Map "General" shapes (Star, Cross, Triangle types).
- [ ] Map "Arrows" library (Block arrows).
- [ ] Map "Flowchart" shapes (Database, Document, Manual Input, etc.).
- [ ] Implement a fallback mechanism that approximates unknown shapes better than a default rectangle (e.g., using the bounding box).

## Issue 2: Advanced Connector Styling
**Type:** Feature
**Priority:** Medium

**Description:**
Connectors currently default to simple black lines with standard arrowheads. Draw.io allows for various line styles (dashed, dotted) and arrowhead types (diamond, circle, none).

**Tasks:**
- [ ] Parse `dashed` and `dashPattern` styles.
- [ ] Map Draw.io start/end arrow styles (`startArrow`, `endArrow`) to PowerPoint line ends (`MSO_LINE_DASH_STYLE`, `MSO_ARROWHEAD_STYLE`).
- [ ] Support "Waypoints" (curved lines vs. sharp elbows).

## Issue 3: Support for Embedded Images
**Type:** Feature
**Priority:** Medium

**Description:**
Diagrams often contain embedded images (logos, screenshots). These are currently ignored by the converter.

**Tasks:**
- [ ] Detect `image` style or `<mxImage>` elements.
- [ ] Extract base64 encoded image data from the XML.
- [ ] Insert the image into the PowerPoint slide at the correct coordinates.

## Issue 4: Grouping Support
**Type:** Feature
**Priority:** Low

**Description:**
Grouped elements in Draw.io currently act as individual items in PowerPoint. Maintaining the group structure would allow easier manipulation in PowerPoint.

**Tasks:**
- [ ] Identify `<mxCell>` items that are parents to others.
- [ ] Use `python-pptx` grouping functionality (if available/reliable) to group the corresponding shapes.

## Issue 5: Rich Text Parsing
**Type:** Enhancement
**Priority:** Medium

**Description:**
HTML labels in Draw.io (e.g., lists, mixed colors in one text box) are stripped to plain text or have very basic formatting.

**Tasks:**
- [ ] Implement a better HTML-to-TextRun parser.
- [ ] Support `<ul>`/`<ol>` lists.
- [ ] Support mixed formatting (e.g., one word bold, next word normal) within the same text box.

## Issue 6: CI/CD Pipeline
**Type:** Infrastructure
**Priority:** Low

**Description:**
Automate testing and linting to ensure code quality.

**Tasks:**
- [ ] Create `.github/workflows/test.yml`.
- [ ] Run `uv run tests/verify_output.py`.
- [ ] Run `ruff` or `flake8` linting.
