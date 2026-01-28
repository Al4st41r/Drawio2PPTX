from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
import re

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

def set_line_end(line, head_type='none', tail_type='none'):
    """
    Manually sets the headEnd and tailEnd of a line using OXML.
    """
    # Ensure ln exists
    ln = line._get_or_add_ln()
    
    # Namespaces
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
    # Helper to set end
    def _set_end(tag_name, val):
        elem = ln.find(f"{{{a_ns}}}{tag_name}")
        if elem is None:
            # Create new
            # We use 'med' for width/len as default
            xml = f'<a:{tag_name} {nsdecls("a")} type="{val}" w="med" len="med"/>'
            elem = parse_xml(xml)
            ln.append(elem)
        else:
            elem.set('type', val)
            
    _set_end('headEnd', head_type)
    _set_end('tailEnd', tail_type)

class HtmlTextParser:
    """Simple parser to convert HTML-like Draw.io strings into segments for PPTX runs."""
    
    def __init__(self, html_text):
        self.raw_text = html_text
        self.segments = []
        
    def parse(self):
        # normalize
        text = self.raw_text.replace('<div>', '\n').replace('</div>', '').replace('<br>', '\n')
        
        parts = re.split(r'(</?[a-zA-Z0-9]+[^>]*>)', text)
        
        current_format = {
            'bold': False,
            'italic': False,
            'underline': False,
            'color': None,
            'size': None
        }
        
        self.segments = []
        
        for part in parts:
            if not part:
                continue
                
            if part.startswith('<'):
                # Tag
                tag = part.lower()
                if tag == '<b>' or 'font-weight: bold' in tag:
                    current_format['bold'] = True
                elif tag == '</b>':
                    current_format['bold'] = False
                elif tag == '<i>' or 'font-style: italic' in tag:
                    current_format['italic'] = True
                elif tag == '</i>':
                    current_format['italic'] = False
                elif tag == '<u>':
                    current_format['underline'] = True
                elif tag == '</u>':
                    current_format['underline'] = False
                elif tag.startswith('<font'):
                    # Extract color
                    m = re.search(r'color="([^"]+)"', part)
                    if m:
                        current_format['color'] = m.group(1)
                elif tag == '</font>':
                    current_format['color'] = None
            else:
                self.segments.append({
                    'text': part,
                    'format': current_format.copy()
                })
        
        return self.segments
