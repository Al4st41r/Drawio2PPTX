from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

from .ppt_map import get_shape_type, get_line_dash, get_arrow_type, get_connector_type
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
                
                conn_type = get_connector_type(e["style"])

                connector = self.slide.shapes.add_connector(
                    conn_type, 0, 0, 0, 0
                )

                # Smart connection logic
                self._connect_shapes(connector, src_shape, tgt_shape, e["style"])

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
            
        tf = shape.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        
        parser = HtmlTextParser(text_value)
        segments = parser.parse()
        
        for seg in segments:
            run = p.add_run()
            run.text = seg['text']
            fmt = seg['format']
            
            run.font.bold = fmt['bold']
            run.font.italic = fmt['italic']
            run.font.underline = fmt['underline']
            
            # Color priority: Segment tag > Style attribute > Default Black
            color_hex = fmt.get('color') or style.get('fontColor') or '#000000'
            rgb = hex_to_rgb(color_hex)
            if rgb:
                run.font.color.rgb = rgb
            
            size = fmt.get('size') or style.get('fontSize')
            if size:
                try:
                    run.font.size = Pt(float(size))
                except:
                    pass

    def _connect_shapes(self, connector, src_shape, tgt_shape, edge_style):
        # Heuristic for connection points (0:Top, 1:Right, 2:Bottom, 3:Left)
        
        def get_idx_from_ratio(x, y):
            """Convert Draw.io 0.0-1.0 ratios to PPTX 0-3 indices."""
            try:
                xf = float(x)
                yf = float(y)
                if yf <= 0.1: return 0 # Top
                if xf >= 0.9: return 1 # Right
                if yf >= 0.9: return 2 # Bottom
                if xf <= 0.1: return 3 # Left
            except:
                pass
            return None

        # Try to get explicit points from style
        src_idx = get_idx_from_ratio(edge_style.get('exitX'), edge_style.get('exitY'))
        tgt_idx = get_idx_from_ratio(edge_style.get('entryX'), edge_style.get('entryY'))

        if src_idx is None or tgt_idx is None:
            # Fallback: Find closest connection points
            def get_site_coords(shape, idx):
                # 0:Top, 1:Right, 2:Bottom, 3:Left
                if idx == 0: return shape.left + shape.width/2, shape.top
                if idx == 1: return shape.left + shape.width, shape.top + shape.height/2
                if idx == 2: return shape.left + shape.width/2, shape.top + shape.height
                if idx == 3: return shape.left, shape.top + shape.height/2
                return 0, 0

            best_dist = float('inf')
            best_pair = (0, 0)
            
            # Check all 4x4 combinations
            # If we already have a fixed src or tgt, we only iterate the other
            src_range = [src_idx] if src_idx is not None else range(4)
            tgt_range = [tgt_idx] if tgt_idx is not None else range(4)
            
            for s in src_range:
                for t in tgt_range:
                    sx, sy = get_site_coords(src_shape, s)
                    tx, ty = get_site_coords(tgt_shape, t)
                    
                    # Euclidean distance
                    dist = ((sx - tx)**2 + (sy - ty)**2)**0.5
                    
                    # Penalize "opposite" connections for Orthogonal lines to prevent overlapping
                    # e.g. Right -> Right is bad
                    if s == t: dist *= 1.5 
                    
                    # Penalize "backwards" flow slightly?
                    # e.g. Right -> Left is usually good (Forward)
                    # Left -> Right is good
                    
                    if dist < best_dist:
                        best_dist = dist
                        best_pair = (s, t)
            
            src_idx, tgt_idx = best_pair
                    
        try:
            connector.begin_connect(src_shape, src_idx)
            connector.end_connect(tgt_shape, tgt_idx)
        except:
            # Fallback to basic connection if index is invalid for shape
            try:
                connector.begin_connect(src_shape, 0)
                connector.end_connect(tgt_shape, 0)
            except:
                pass


    def _add_edge_label(self, edge, src_shape, tgt_shape):
        # Calculate midpoint
        mid_x = (src_shape.left + tgt_shape.left) / 2
        mid_y = (src_shape.top + tgt_shape.top) / 2
        
        # Offset adjustment
        geo = edge["geometry"]
        offset_x = 0
        offset_y = 0
        
        if geo is not None:
            # mxGeometry for edge label offset
            try:
                # y in mxGeometry for relative edge labels is often used as the offset
                # We check relative="1" usually but Draw.io handles it specifically
                offset_y = float(geo.get('y', 0))
                offset_x = float(geo.get('x', 0))
            except:
                pass
            
            # Explicit offset point (user dragged label)
            offset = geo.find("mxPoint")
            if offset is not None and offset.get("as") == "offset":
                try:
                    ox = float(offset.get("x", 0))
                    oy = float(offset.get("y", 0))
                    if ox != 0 or oy != 0:
                        offset_x = ox
                        offset_y = oy
                except:
                    pass

        # Create a text box
        # Labels should be large enough to not wrap single words
        box_w = px_to_emu(80)
        box_h = px_to_emu(40)
        
        left = mid_x + px_to_emu(offset_x) - (box_w / 2)
        top = mid_y + px_to_emu(offset_y) - (box_h / 2)

        tb = self.slide.shapes.add_textbox(left, top, box_w, box_h)
        tb.text_frame.word_wrap = False 
        
        # Style the textbox background (optional, Draw.io labels are usually transparent or white)
        # For now, let's keep it transparent but ensure text is visible
        
        self._apply_text(tb, edge["value"], edge["style"])
