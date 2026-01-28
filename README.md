# Drawio2PPTX

A Python-based tool to convert Draw.io (`.drawio`, `.xml`) diagrams into editable Microsoft PowerPoint (`.pptx`) presentations.

## Features

-   **Editable Shapes:** Converts rectangles, diamonds, ellipses, etc. to native PowerPoint shapes.
-   **Dynamic Connectors:** Arrows are real connectors that stick to shapes when moved.
-   **Text Styling:** Preserves font size, bold, italic, underline, and colors.
-   **Layout:** Accurate positioning and sizing.

## Installation

This project uses `uv` for dependency management.

```bash
uv sync
```

## Usage (CLI)

```bash
uv run main.py input.drawio output.pptx
```

## Usage (Web App)

A Flask-based web interface is included.

### Development
```bash
uv run webapp/app.py
```
Access at `http://localhost:5003`.

### Deployment (Systemd + Nginx)

1.  **Service Setup:**
    ```bash
    sudo cp webapp/drawio2pptx.service /etc/systemd/system/
    sudo systemctl daemon-reload
    sudo systemctl enable drawio2pptx
    sudo systemctl start drawio2pptx
    ```

2.  **Nginx Setup:**
    Add the contents of `webapp/nginx-config.txt` to your Nginx site configuration.

## Project Structure

-   `main.py`: CLI entry point.
-   `converter/`: Core conversion logic.
    -   `parser.py`: Parses Draw.io XML.
    -   `engine.py`: Generates PPTX using `python-pptx`.
    -   `ppt_map.py`: Maps Draw.io shapes to PowerPoint shapes.
-   `webapp/`: Flask web application.
-   `tests/`: Unit tests and verification scripts.
