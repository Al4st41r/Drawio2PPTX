from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE

# Mapping from Draw.io shape names/styles to python-pptx MSO_SHAPE
SHAPE_MAP = {
    'rectangle': MSO_SHAPE.RECTANGLE,
    'rounded': MSO_SHAPE.ROUNDED_RECTANGLE,
    'ellipse': MSO_SHAPE.OVAL,
    'rhombus': MSO_SHAPE.DIAMOND,
    'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'hexagon': MSO_SHAPE.HEXAGON,
    'cloud': MSO_SHAPE.CLOUD,
    'cylinder': MSO_SHAPE.CAN,
    'actor': MSO_SHAPE.SMILEY_FACE, # Approximation
    'process': MSO_SHAPE.RECTANGLE, 
    'decision': MSO_SHAPE.DIAMOND, 
    'note': MSO_SHAPE.FOLDED_CORNER,
    'callout': MSO_SHAPE.RECTANGULAR_CALLOUT,
    'step': MSO_SHAPE.RECTANGLE,
    'parallelogram': MSO_SHAPE.PARALLELOGRAM,
}

LINE_DASH_MAP = {
    '1': MSO_LINE_DASH_STYLE.DASH,
    'dashed': MSO_LINE_DASH_STYLE.DASH,
    'dotted': MSO_LINE_DASH_STYLE.ROUND_DOT,
    'dashDot': MSO_LINE_DASH_STYLE.DASH_DOT
}

# Mapping Draw.io arrow style names to PPTX XML type strings
ARROW_MAP = {
    'none': 'none',
    'block': 'triangle',
    'classic': 'triangle',
    'open': 'arrow',
    'oval': 'oval',
    'diamond': 'diamond',
    'thindiamond': 'diamond',
    'erMany': 'triangle',
    'erOne': 'stealth',
    'dash': 'none',
    'standard': 'triangle',
    'async': 'arrow'
}

def get_shape_type(style_dict):
    """Determines the best MSO_SHAPE type based on the style dictionary."""
    if 'shape' in style_dict:
        shape_name = style_dict['shape']
        if shape_name in SHAPE_MAP:
            return SHAPE_MAP[shape_name]
            
    for key in style_dict:
        if key in SHAPE_MAP:
            return SHAPE_MAP[key]
            
    if 'rounded' in style_dict and style_dict['rounded'] == '1':
        return MSO_SHAPE.ROUNDED_RECTANGLE
    if 'ellipse' in style_dict:
        return MSO_SHAPE.OVAL
        
    return MSO_SHAPE.RECTANGLE

def get_line_dash(style_dict):
    if 'dashed' in style_dict and style_dict['dashed'] == '1':
        return MSO_LINE_DASH_STYLE.DASH
    return MSO_LINE_DASH_STYLE.SOLID

def get_arrow_type(arrow_style, fill=True):
    # If not filled, some styles look better as 'arrow' (open)
    if not fill and arrow_style in ['block', 'classic']:
        return 'arrow'
    return ARROW_MAP.get(arrow_style, 'none')

def get_arrow_size(size_str):
    """Maps endSize to w/len strings."""
    try:
        size = float(size_str)
        if size < 6: return 'sm', 'sm'
        if size > 12: return 'lg', 'lg'
    except:
        pass
    return 'med', 'med'

def get_connector_type(style_dict):
    """Maps Draw.io edge styles to PPTX connector types."""
    if 'curved' in style_dict and style_dict['curved'] == '1':
        return MSO_CONNECTOR.CURVE
    
    edge_style = style_dict.get('edgeStyle', '')
    if 'orthogonal' in edge_style or 'elbow' in edge_style:
        return MSO_CONNECTOR.ELBOW
    
    if not edge_style or edge_style == 'none':
        return MSO_CONNECTOR.STRAIGHT
        
    return MSO_CONNECTOR.ELBOW # Default to elbow for safety
