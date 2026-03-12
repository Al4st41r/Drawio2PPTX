import re

from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls


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

from html.parser import HTMLParser as _StdHTMLParser


class HtmlTextParser:
    """Parses Draw.io HTML labels into structured paragraphs with formatted runs.

    Returns a list of paragraph dicts, each containing:
        - "runs": list of {"text", "bold", "italic", "underline", "color", "size"}
        - "indent_level": int (0 for normal text, 1+ for nested lists)
        - "bullet": None | "bullet" | "number"
    """

    def __init__(self, html_text):
        self.raw_text = html_text

    def parse(self):
        parser = _HtmlToParas()
        parser.feed(self.raw_text)
        paragraphs = parser.get_paragraphs()

        if not paragraphs:
            # Plain text fallback (no HTML tags at all)
            clean = re.sub(r'<[^>]+>', '', self.raw_text).strip()
            if clean:
                paragraphs = [_make_para(runs=[_make_run(text=clean)])]

        return paragraphs


# Block-level tags that force a new paragraph
_BLOCK_TAGS = {'div', 'p', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}

_DEFAULT_FMT = {
    'bold': False,
    'italic': False,
    'underline': False,
    'color': None,
    'size': None,
}


def _make_run(text='', bold=False, italic=False, underline=False,
              color=None, size=None):
    return {
        'text': text,
        'bold': bold,
        'italic': italic,
        'underline': underline,
        'color': color,
        'size': size,
    }


def _make_para(runs=None, indent_level=0, bullet=None):
    return {
        'runs': runs or [],
        'indent_level': indent_level,
        'bullet': bullet,
    }


class _HtmlToParas(_StdHTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=False)
        self._paragraphs = []
        self._cur_para = _make_para()
        self._fmt_stack = [dict(_DEFAULT_FMT)]
        # List state: stack of ("ul"|"ol", counter)
        self._list_stack = []
        self._in_li = False

    # -- helpers --

    def _cur_fmt(self):
        return self._fmt_stack[-1]

    def _push_fmt(self, **overrides):
        new = dict(self._cur_fmt())
        new.update(overrides)
        self._fmt_stack.append(new)

    def _pop_fmt(self):
        if len(self._fmt_stack) > 1:
            self._fmt_stack.pop()

    def _start_new_para(self, bullet=None):
        # Flush current paragraph if it has content
        if self._cur_para['runs']:
            self._paragraphs.append(self._cur_para)
        indent = len(self._list_stack)
        self._cur_para = _make_para(indent_level=indent, bullet=bullet)

    def _add_text(self, text):
        if not text:
            return
        fmt = self._cur_fmt()
        self._cur_para['runs'].append(_make_run(
            text=text,
            bold=fmt['bold'],
            italic=fmt['italic'],
            underline=fmt['underline'],
            color=fmt['color'],
            size=fmt['size'],
        ))

    # -- extract inline style attributes --

    @staticmethod
    def _extract_attrs_from_tag(attrs):
        """Pull colour, size, bold, italic from tag attributes and inline style."""
        overrides = {}
        attr_dict = dict(attrs)

        # <font color="..."> or <font size="...">
        if 'color' in attr_dict:
            overrides['color'] = attr_dict['color']
        if 'size' in attr_dict:
            overrides['size'] = attr_dict['size']

        style = attr_dict.get('style', '')
        if style:
            m = re.search(r'(?:^|;|\s)color\s*:\s*([^;]+)', style, re.I)
            if m:
                overrides['color'] = m.group(1).strip().strip("'\"")
            m = re.search(r'font-size\s*:\s*([^;]+)', style, re.I)
            if m:
                raw = m.group(1).strip().strip("'\"")
                # Strip 'px' / 'pt' suffix to get a numeric value
                cleaned = re.sub(r'\s*(px|pt)\s*$', '', raw, flags=re.I)
                overrides['size'] = cleaned
            if re.search(r'font-weight\s*:\s*bold', style, re.I):
                overrides['bold'] = True
            if re.search(r'font-style\s*:\s*italic', style, re.I):
                overrides['italic'] = True
            if re.search(r'text-decoration[^:]*:\s*underline', style, re.I):
                overrides['underline'] = True

        return overrides

    # -- HTMLParser callbacks --

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()

        if tag == 'br':
            self._start_new_para()
            return

        if tag in ('div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self._start_new_para()
            overrides = self._extract_attrs_from_tag(attrs)
            if tag.startswith('h'):
                overrides.setdefault('bold', True)
            self._push_fmt(**overrides)
            return

        if tag == 'ul':
            self._list_stack.append(('ul', 0))
            return
        if tag == 'ol':
            self._list_stack.append(('ol', 0))
            return

        if tag == 'li':
            self._in_li = True
            if self._list_stack:
                list_type, count = self._list_stack[-1]
                count += 1
                self._list_stack[-1] = (list_type, count)
                if list_type == 'ul':
                    self._start_new_para(bullet='bullet')
                else:
                    self._start_new_para(bullet='number')
                    self._add_text(f'{count}. ')
            else:
                # <li> outside a list — treat as bullet
                self._start_new_para(bullet='bullet')
            return

        if tag in ('b', 'strong'):
            self._push_fmt(bold=True)
            return
        if tag in ('i', 'em'):
            self._push_fmt(italic=True)
            return
        if tag == 'u':
            self._push_fmt(underline=True)
            return

        if tag in ('font', 'span'):
            overrides = self._extract_attrs_from_tag(attrs)
            self._push_fmt(**overrides)
            return

    def handle_endtag(self, tag):
        tag = tag.lower()

        if tag == 'br':
            return

        if tag in ('div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self._pop_fmt()
            self._start_new_para()
            return

        if tag in ('ul', 'ol'):
            if self._list_stack:
                self._list_stack.pop()
            return

        if tag == 'li':
            self._in_li = False
            return

        if tag in ('b', 'strong', 'i', 'em', 'u', 'font', 'span'):
            self._pop_fmt()
            return

    def handle_data(self, data):
        self._add_text(data)

    def handle_entityref(self, name):
        from html import unescape
        self._add_text(unescape(f'&{name};'))

    def handle_charref(self, name):
        from html import unescape
        self._add_text(unescape(f'&#{name};'))

    def get_paragraphs(self):
        # Flush last paragraph
        if self._cur_para['runs']:
            self._paragraphs.append(self._cur_para)

        # Strip leading/trailing empty-text-only runs and empty paragraphs
        result = []
        for para in self._paragraphs:
            # Collapse consecutive whitespace-only runs but keep structure
            cleaned_runs = []
            for run in para['runs']:
                text = run['text']
                if text == '\n' and not cleaned_runs:
                    continue  # skip leading newlines within a paragraph
                cleaned_runs.append(run)
            # Trim trailing whitespace-only runs
            while cleaned_runs and cleaned_runs[-1]['text'].strip() == '':
                cleaned_runs.pop()
            if cleaned_runs:
                para['runs'] = cleaned_runs
                result.append(para)

        # Add bullet prefix for unordered list items
        for para in result:
            if para['bullet'] == 'bullet':
                if para['runs'] and not para['runs'][0]['text'].startswith('\u2022'):
                    para['runs'][0] = dict(para['runs'][0])
                    para['runs'][0]['text'] = '\u2022 ' + para['runs'][0]['text']

        return result
