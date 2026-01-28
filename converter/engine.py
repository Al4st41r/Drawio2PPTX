from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

from .ppt_map import get_shape_type, get_line_dash
from .utils import hex_to_rgb, px_to_emu, HtmlTextParser


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

    def _apply_text(self, shape, text_value, style):
        if not text_value:
            return
            
        # Clear default paragraph
        if not shape.text_frame.text.strip():
            shape.text_frame.clear()
        
        # Parse HTML text
        parser = HtmlTextParser(text_value)
        segments = parser.parse()
        
        # Use first paragraph or create new if needed
        p = shape.text_frame.paragraphs[0]
        
        for seg in segments:
            run = p.add_run()
            run.text = seg['text']
            fmt = seg['format']
            
            if fmt['bold']:
                run.font.bold = True
            if fmt['italic']:
                run.font.italic = True
            if fmt['underline']:
                run.font.underline = True
            if fmt['color']:
                rgb = hex_to_rgb(fmt['color'])
                if rgb:
                    run.font.color.rgb = rgb
            
            # Global style fallback
            if 'fontSize' in style:
                try:
                    run.font.size = Pt(float(style['fontSize']))
                except:
                    pass

    def _connect_shapes(self, connector, src_shape, tgt_shape):
        src_x = src_shape.left + src_shape.width / 2
        src_y = src_shape.top + src_shape.height / 2
        tgt_x = tgt_shape.left + tgt_shape.width / 2
        tgt_y = tgt_shape.top + tgt_shape.height / 2

        src_idx = 2  # Default bottom
        tgt_idx = 0  # Default top

        # Simple heuristic for connection points (0-3)
        if abs(src_x - tgt_x) > abs(src_y - tgt_y):
            if src_x < tgt_x:  # Target is right
                src_idx = 1
                tgt_idx = 3
            else:  # Target is left
                src_idx = 3
                tgt_idx = 1
        else:
            if src_y < tgt_y:  # Target is below
                src_idx = 2
                tgt_idx = 0
            else:  # Target is above
                src_idx = 0
                tgt_idx = 2

        connector.begin_connect(src_shape, src_idx)
        connector.end_connect(tgt_shape, tgt_idx)

    def _add_edge_label(self, edge, src_shape, tgt_shape):
        # Calculate midpoint
        mid_x = (src_shape.left + tgt_shape.left) / 2
        mid_y = (src_shape.top + tgt_shape.top) / 2

        # Offset adjustment if provided in geometry (mxPoint as offset)
        geo = edge["geometry"]
        offset_x = 0
        offset_y = 0
        if geo is not None:
            # mxGeometry for edge often has x,y as relative points,
            # but <mxPoint as="offset" /> handles label offset
            offset = geo.find("mxPoint")
            if offset is not None and offset.get("as") == "offset":
                offset_x = float(offset.get("x", 0))
                offset_y = float(offset.get("y", 0))

        # Create a text box
        # We need to approximate position in EMUs.
        # Draw.io offsets are pixels.

        left = mid_x + px_to_emu(offset_x) - px_to_emu(20)  # Center roughly
        top = mid_y + px_to_emu(offset_y) - px_to_emu(10)

        tb = self.slide.shapes.add_textbox(
            left, top, px_to_emu(40), px_to_emu(20))
        self._apply_text(tb, edge["value"], edge["style"])
