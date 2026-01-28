from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
import re

def px_to_emu(px):
    return int(float(px) * 914400 / 96)

def hex_to_rgb(hex_color):
    if not hex_color or hex_color == 'none':
        return None
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 3:
        hex_color = ''.join([c*2 for c in hex_color])
    if len(hex_color) != 6:
        return None
    try:
        return RGBColor(int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:], 16))
    except ValueError:
        return None

def parse_style_string(style_str):
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

def set_line_end(line, head_type='none', tail_type='none', head_w='med', head_l='med', tail_w='med', tail_l='med'):
    ln = line._get_or_add_ln()
    a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    
    def _set_end(tag_name, val, w, l):
        elem = ln.find(f"{{{a_ns}}}{tag_name}")
        if elem is None:
            xml = f'<a:{tag_name} {nsdecls("a")} type="{val}" w="{w}" len="{l}"/>'
            elem = parse_xml(xml)
            ln.append(elem)
        else:
            elem.set('type', val)
            elem.set('w', w)
            elem.set('len', l)
            
    _set_end('headEnd', head_type, head_w, head_l)
    _set_end('tailEnd', tail_type, tail_w, tail_l)

class HtmlTextParser:
    def __init__(self, html_text):
        self.raw_text = html_text
        
    def parse(self):
        # normalize
        text = self.raw_text.replace('<div>', '\n').replace('</div>', '\n').replace('<br>', '\n')
        
        # Strip outer containers if they are just spans/divs without style
        # But we'll just split by any tag
        parts = re.split(r'(</?[a-zA-Z0-9]+[^>]*>)', text)
        
        # Default state
        current_format = {
            'bold': False,
            'italic': False,
            'underline': False,
            'color': '#000000', # Default to black
            'size': None
        }
        
        segments = []
        for part in parts:
            if not part: continue
            if part.startswith('<'):
                tag = part.lower()
                if '<b>' in tag or 'font-weight: bold' in tag or 'font-weight:bold' in tag:
                    current_format['bold'] = True
                elif '</b>' in tag:
                    current_format['bold'] = False
                elif '<i>' in tag or 'font-style: italic' in tag:
                    current_format['italic'] = True
                elif '</i>' in tag:
                    current_format['italic'] = False
                elif '<u>' in tag:
                    current_format['underline'] = True
                elif '</u>' in tag:
                    current_format['underline'] = False
                elif tag.startswith('<font') or tag.startswith('<span'):
                    # Extract color
                    m = re.search(r'color[:=]\s*["\']?([^"\';\s>]+)["\']?', part, re.I)
                    if m: current_format['color'] = m.group(1)
                    # Extract size
                    s = re.search(r'size[:=]\s*["\']?([^"\';\s>]+)["\']?', part, re.I)
                    if s: current_format['size'] = s.group(1)
                elif tag == '</font>' or tag == '</span>':
                    # Resetting is hard without a stack, but Draw.io usually doesn't nest deeply
                    # For now we'll just keep the last set color/size if nested
                    pass
            else:
                segments.append({'text': part, 'format': current_format.copy()})
        
        if not segments and text:
            # Fallback if no tags but text exists
            segments.append({'text': text, 'format': current_format.copy()})
            
        return segments