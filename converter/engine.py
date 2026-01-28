from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

from .ppt_map import get_shape_type, get_line_dash, get_arrow_type
from .utils import hex_to_rgb, px_to_emu, HtmlTextParser, set_line_end


class PptxGenerator:
    def __init__(self, output_file):
        self.output_file = output_file
        self.prs = Presentation()
        self.slide = None
        self.id_to_shape = {}  # Map Draw.io ID to PPTX Shape per slide

    def create_slide(self):
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.id_to_shape = {} # Reset map for new slide
        return self.slide

    def add_vertices(self, vertices):
        if not self.slide:
            self.create_slide()
            
        for v in vertices:
            x, y, w, h = v["x"], v["y"], v["width"], v["height"]
            style = v["style"]

            shape_type = get_shape_type(style)

            shape = self.slide.shapes.add_shape(
                shape_type, px_to_emu(x), px_to_emu(
                    y), px_to_emu(w), px_to_emu(h)
            )

            # Apply Styles
            self._apply_shape_style(shape, style)

            # Add Text
            if v["value"]:
                self._apply_text(shape, v["value"], style)

            self.id_to_shape[v["id"]] = shape

    def add_edges(self, edges):
        for e in edges:
            source_id = e["source"]
            target_id = e["target"]

            if source_id in self.id_to_shape and target_id in self.id_to_shape:
                src_shape = self.id_to_shape[source_id]
                tgt_shape = self.id_to_shape[target_id]

                connector = self.slide.shapes.add_connector(
                    MSO_CONNECTOR.ELBOW, 0, 0, 0, 0
                )

                # Smart connection logic
                self._connect_shapes(connector, src_shape, tgt_shape)

                # Apply Styles
                self._apply_line_style(connector.line, e["style"])
                
                # Edge Label (if any)
                if e['value']:
                    self._add_edge_label(e, src_shape, tgt_shape)

    def save(self):
        self.prs.save(self.output_file)

    def _apply_shape_style(self, shape, style):
        # Fill
        fill_color = style.get("fillColor")
        if fill_color == "none":
            shape.fill.background()  # No fill
        elif fill_color:
            rgb = hex_to_rgb(fill_color)
            if rgb:
                shape.fill.solid()
                shape.fill.fore_color.rgb = rgb
        else:
            # Default white fill if not specified
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_to_rgb("#FFFFFF")

        # Line / Stroke
        stroke_color = style.get("strokeColor")
        if stroke_color == "none":
            shape.line.fill.background()
        elif stroke_color:
            rgb = hex_to_rgb(stroke_color)
            if rgb:
                shape.line.color.rgb = rgb
        else:
            # Default black stroke
            shape.line.color.rgb = hex_to_rgb("#000000")

        # Stroke Width
        if "strokeWidth" in style:
            try:
                width_pt = float(style["strokeWidth"])
                shape.line.width = Pt(width_pt)
            except:
                pass

    def _apply_line_style(self, line, style):
        # Color
        stroke_color = style.get("strokeColor")
        if stroke_color and stroke_color != "none":
            rgb = hex_to_rgb(stroke_color)
            if rgb:
                line.color.rgb = rgb
        else:
            line.color.rgb = hex_to_rgb("#000000")  # Default black

        # Width
        if "strokeWidth" in style:
            try:
                width_pt = float(style["strokeWidth"])
                line.width = Pt(width_pt)
            except:
                pass
        else:
            line.width = Pt(1)
            
        # Dash Style
        line.dash_style = get_line_dash(style)
        
        # Arrowheads (Manual XML injection)
        from .ppt_map import get_arrow_type, get_arrow_size
        
        start_fill = style.get('startFill') != '0'
        end_fill = style.get('endFill') != '0'
        
        start_arrow = get_arrow_type(style.get('startArrow', 'none'), fill=start_fill)
        end_arrow = get_arrow_type(style.get('endArrow', 'none'), fill=end_fill)
        
        # Size mapping
        start_w, start_l = get_arrow_size(style.get('startSize', '6'))
        end_w, end_l = get_arrow_size(style.get('endSize', '6'))
        
        set_line_end(line, head_type=end_arrow, tail_type=start_arrow, 
                     head_w=end_w, head_l=end_l, tail_w=start_w, tail_l=start_l)

    def _apply_text(self, shape, text_value, style):
        if not text_value:
            return
            
        # Clear default paragraph
        if not shape.text_frame.text.strip():
            shape.text_frame.clear()
        
        parser = HtmlTextParser(text_value)
        segments = parser.parse()
        
        p = shape.text_frame.paragraphs[0]
        
        for seg in segments:
            run = p.add_run()
            run.text = seg['text']
            fmt = seg['format']
            
            if fmt['bold']: run.font.bold = True
            if fmt['italic']: run.font.italic = True
            if fmt['underline']: run.font.underline = True
            if fmt['color']:
                rgb = hex_to_rgb(fmt['color'])
                if rgb: run.font.color.rgb = rgb
            
            if fmt['size']:
                try:
                    run.font.size = Pt(float(fmt['size']))
                except: pass
            elif 'fontSize' in style:
                try:
                    run.font.size = Pt(float(style['fontSize']))
                except: pass

    def _connect_shapes(self, connector, src_shape, tgt_shape):
        # We need to find the mxCell for this edge to get exitX/Y etc.
        # But for now let's just use the shapes' relative positions.
        
        src_x, src_y = src_shape.left, src_shape.top
        src_w, src_h = src_shape.width, src_shape.height
        tgt_x, tgt_y = tgt_shape.left, tgt_shape.top
        tgt_w, tgt_h = tgt_shape.width, tgt_shape.height
        
        # Center points
        scx, scy = src_x + src_w/2, src_y + src_h/2
        tcx, tcy = tgt_x + tgt_w/2, tgt_y + tgt_h/2
        
        # Heuristic for connection points (0:Top, 1:Right, 2:Bottom, 3:Left)
        # TODO: Read exitX/Y from Draw.io
        
        if abs(scx - tcx) > abs(scy - tcy):
            # Horizontal dominant
            if scx < tcx: src_idx = 1; tgt_idx = 3 # Right -> Left
            else: src_idx = 3; tgt_idx = 1 # Left -> Right
        else:
            # Vertical dominant
            if scy < tcy: src_idx = 2; tgt_idx = 0 # Bottom -> Top
            else: src_idx = 0; tgt_idx = 2 # Top -> Bottom
            
        try:
            connector.begin_connect(src_shape, src_idx)
            connector.end_connect(tgt_shape, tgt_idx)
        except:
            pass # Sometimes shapes don't support certain indices


    def _add_edge_label(self, edge, src_shape, tgt_shape):
        # Calculate midpoint
        mid_x = (src_shape.left + tgt_shape.left) / 2
        mid_y = (src_shape.top + tgt_shape.top) / 2
        
        # Offset adjustment
        geo = edge["geometry"]
        offset_x = 0
        offset_y = 0
        
        if geo is not None:
            # Check geometry attributes first (often used for relative offsets)
            if geo.get('relative') == '1':
                try:
                    offset_x = float(geo.get('x', 0))
                    offset_y = float(geo.get('y', 0))
                except ValueError:
                    pass
            
            # Check child offset point (overrides or adds?)
            # Usually it's either/or. Draw.io uses mxPoint as="offset" for label position
            offset = geo.find("mxPoint")
            if offset is not None and offset.get("as") == "offset":
                try:
                    ox = float(offset.get("x", 0))
                    oy = float(offset.get("y", 0))
                    if ox != 0 or oy != 0:
                        offset_x = ox
                        offset_y = oy
                except ValueError:
                    pass

        # Create a text box
        # We need to approximate position in EMUs.
        # Draw.io offsets are pixels.
        
        # PPTX text boxes are top-left based. We want center based.
        box_w = px_to_emu(40)
        box_h = px_to_emu(20)
        
        left = mid_x + px_to_emu(offset_x) - (box_w / 2)
        top = mid_y + px_to_emu(offset_y) - (box_h / 2)

        tb = self.slide.shapes.add_textbox(left, top, box_w, box_h)
        self._apply_text(tb, edge["value"], edge["style"])
        
        # Center text in the label box
        tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
