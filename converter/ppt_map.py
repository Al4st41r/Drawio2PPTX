from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

# Mapping from Draw.io shape names/styles to python-pptx MSO_SHAPE
# This is a heuristic mapping.
SHAPE_MAP = {
    'rectangle': MSO_SHAPE.RECTANGLE,
    'rounded': MSO_SHAPE.ROUNDED_RECTANGLE,
    'ellipse': MSO_SHAPE.OVAL,
    'rhombus': MSO_SHAPE.DIAMOND,
    'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'hexagon': MSO_SHAPE.HEXAGON,
    'cloud': MSO_SHAPE.CLOUD,
    'cylinder': MSO_SHAPE.CAN,
    'actor': MSO_SHAPE.SMILEY_FACE, # Approximation, usually stick figure
    'process': MSO_SHAPE.RECTANGLE, # Flowchart process
    'decision': MSO_SHAPE.DIAMOND, # Flowchart decision
    'note': MSO_SHAPE.FOLDED_CORNER, # Sticky note
}

def get_shape_type(style_dict):
    """Determines the best MSO_SHAPE type based on the style dictionary."""
    
    # Check for specific shape keys
    if 'shape' in style_dict:
        shape_name = style_dict['shape']
        if shape_name in SHAPE_MAP:
            return SHAPE_MAP[shape_name]
            
    # Check for loose keys (e.g. style="rhombus;whiteSpace=wrap...")
    for key in style_dict:
        if key in SHAPE_MAP:
            return SHAPE_MAP[key]
            
    # Specific attributes
    if 'rounded' in style_dict and style_dict['rounded'] == '1':
        return MSO_SHAPE.ROUNDED_RECTANGLE
    if 'ellipse' in style_dict:
        return MSO_SHAPE.OVAL
        
    # Default
    return MSO_SHAPE.RECTANGLE
