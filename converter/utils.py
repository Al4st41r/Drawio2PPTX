from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

def px_to_emu(px):
    # 1 inch = 914400 EMUs
    # Assume 96 DPI for screen pixels
    return int(float(px) * 914400 / 96)

def hex_to_rgb(hex_color):
    """Converts hex string '#RRGGBB' to RGBColor object. Returns None if invalid or 'none'."""
    if not hex_color or hex_color == 'none':
        return None
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        return None
    try:
        return RGBColor(int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:], 16))
    except ValueError:
        return None

def parse_style_string(style_str):
    """Parses the style string 'key=value;key2=value2' into a dict."""
    style = {}
    if not style_str:
        return style
    parts = style_str.split(';')
    for part in parts:
        if '=' in part:
            k, v = part.split('=', 1)
            style[k] = v
        elif part:
            style[part] = True
    return style
