import logging

from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE

log = logging.getLogger(__name__)

# ── Core / General shapes ───────────────────────────────────────────────
SHAPE_MAP = {
    # Basic
    'rectangle': MSO_SHAPE.RECTANGLE,
    'rounded': MSO_SHAPE.ROUNDED_RECTANGLE,
    'ellipse': MSO_SHAPE.OVAL,
    'rhombus': MSO_SHAPE.DIAMOND,
    'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'hexagon': MSO_SHAPE.HEXAGON,
    'cloud': MSO_SHAPE.CLOUD,
    'cylinder': MSO_SHAPE.CAN,
    'actor': MSO_SHAPE.SMILEY_FACE,
    'process': MSO_SHAPE.RECTANGLE,
    'decision': MSO_SHAPE.DIAMOND,
    'note': MSO_SHAPE.FOLDED_CORNER,
    'callout': MSO_SHAPE.RECTANGULAR_CALLOUT,
    'step': MSO_SHAPE.RECTANGLE,
    'parallelogram': MSO_SHAPE.PARALLELOGRAM,

    # General – Stars
    'mxgraph.basic.star': MSO_SHAPE.STAR_5_POINT,
    'mxgraph.basic.4pointstar2': MSO_SHAPE.STAR_4_POINT,
    'mxgraph.basic.6pointstar2': MSO_SHAPE.STAR_6_POINT,
    'mxgraph.basic.8pointstar': MSO_SHAPE.STAR_8_POINT,

    # General – Cross / Plus
    'cross': MSO_SHAPE.CROSS,
    'mxgraph.basic.cross2': MSO_SHAPE.CROSS,
    'plus': MSO_SHAPE.CROSS,

    # General – Triangles
    'mxgraph.basic.acute_triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'mxgraph.basic.orthogonal_triangle': MSO_SHAPE.RIGHT_TRIANGLE,

    # General – Other polygons
    'trapezoid': MSO_SHAPE.TRAPEZOID,
    'pentagon': MSO_SHAPE.REGULAR_PENTAGON,
    'mxgraph.basic.pentagon': MSO_SHAPE.REGULAR_PENTAGON,
    'octagon': MSO_SHAPE.OCTAGON,
    'mxgraph.basic.octagon2': MSO_SHAPE.OCTAGON,
    'heptagon': MSO_SHAPE.HEPTAGON,
    'decagon': MSO_SHAPE.DECAGON,
    'dodecagon': MSO_SHAPE.DODECAGON,

    # General – Misc
    'cube': MSO_SHAPE.CUBE,
    'heart': MSO_SHAPE.HEART,
    'mxgraph.basic.heart': MSO_SHAPE.HEART,
    'moon': MSO_SHAPE.MOON,
    'mxgraph.basic.moon': MSO_SHAPE.MOON,
    'arc': MSO_SHAPE.ARC,
    'donut': MSO_SHAPE.DONUT,
    'mxgraph.basic.donut': MSO_SHAPE.DONUT,
    'sun': MSO_SHAPE.SUN,
    'plaque': MSO_SHAPE.PLAQUE,
    'frame': MSO_SHAPE.FRAME,
    'mxgraph.basic.frame': MSO_SHAPE.FRAME,
    'corner': MSO_SHAPE.CORNER,
    'mxgraph.basic.frame_corner': MSO_SHAPE.CORNER,
    'teardrop': MSO_SHAPE.TEAR,
    'mxgraph.basic.teardrop': MSO_SHAPE.TEAR,
    'lightningBolt': MSO_SHAPE.LIGHTNING_BOLT,
    'mxgraph.basic.no_symbol': MSO_SHAPE.NO_SYMBOL,
    'wave': MSO_SHAPE.WAVE,
    'doubleWave': MSO_SHAPE.DOUBLE_WAVE,
    'bevel': MSO_SHAPE.BEVEL,
    'pie': MSO_SHAPE.PIE,
    'funnel': MSO_SHAPE.FUNNEL,
    'mxgraph.basic.funnel': MSO_SHAPE.FUNNEL,
    'smileyFace': MSO_SHAPE.SMILEY_FACE,

    # Braces / Brackets
    'curlyBracket': MSO_SHAPE.LEFT_BRACE,
    'leftBrace': MSO_SHAPE.LEFT_BRACE,
    'rightBrace': MSO_SHAPE.RIGHT_BRACE,
    'leftBracket': MSO_SHAPE.LEFT_BRACKET,
    'rightBracket': MSO_SHAPE.RIGHT_BRACKET,
    'bracePair': MSO_SHAPE.DOUBLE_BRACE,
    'bracketPair': MSO_SHAPE.DOUBLE_BRACKET,

    # Scrolls / Ribbons
    'horizontalScroll': MSO_SHAPE.HORIZONTAL_SCROLL,
    'verticalScroll': MSO_SHAPE.VERTICAL_SCROLL,
    'ribbon': MSO_SHAPE.DOWN_RIBBON,
    'ribbon2': MSO_SHAPE.UP_RIBBON,

    # Callouts
    'mxgraph.basic.rounded_frame': MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    'mxgraph.basic.oval_callout': MSO_SHAPE.OVAL_CALLOUT,
    'mxgraph.basic.cloud_callout': MSO_SHAPE.CLOUD_CALLOUT,
    'wedgeRoundRectCallout': MSO_SHAPE.BALLOON,
    'wedgeEllipseCallout': MSO_SHAPE.OVAL_CALLOUT,

    # ── Arrows library (Block arrows) ──────────────────────────────────
    'mxgraph.arrows2.arrow': MSO_SHAPE.RIGHT_ARROW,
    'mxgraph.arrows2.arrowLeft': MSO_SHAPE.LEFT_ARROW,
    'mxgraph.arrows2.arrowUp': MSO_SHAPE.UP_ARROW,
    'mxgraph.arrows2.arrowDown': MSO_SHAPE.DOWN_ARROW,

    'singleArrow': MSO_SHAPE.RIGHT_ARROW,
    'doubleArrow': MSO_SHAPE.LEFT_RIGHT_ARROW,
    'leftArrow': MSO_SHAPE.LEFT_ARROW,
    'rightArrow': MSO_SHAPE.RIGHT_ARROW,
    'upArrow': MSO_SHAPE.UP_ARROW,
    'downArrow': MSO_SHAPE.DOWN_ARROW,
    'leftRightArrow': MSO_SHAPE.LEFT_RIGHT_ARROW,
    'upDownArrow': MSO_SHAPE.UP_DOWN_ARROW,
    'quadArrow': MSO_SHAPE.QUAD_ARROW,
    'leftUpArrow': MSO_SHAPE.LEFT_UP_ARROW,
    'bentArrow': MSO_SHAPE.BENT_ARROW,
    'bentUpArrow': MSO_SHAPE.BENT_UP_ARROW,
    'uturnArrow': MSO_SHAPE.U_TURN_ARROW,
    'stripedRightArrow': MSO_SHAPE.STRIPED_RIGHT_ARROW,
    'notchedRightArrow': MSO_SHAPE.NOTCHED_RIGHT_ARROW,
    'chevron': MSO_SHAPE.CHEVRON,
    'homePlate': MSO_SHAPE.PENTAGON,
    'circularArrow': MSO_SHAPE.CIRCULAR_ARROW,

    'mxgraph.arrows2.bendDoubleArrow': MSO_SHAPE.BENT_UP_ARROW,
    'mxgraph.arrows2.chevronArrow': MSO_SHAPE.CHEVRON,
    'mxgraph.arrows2.quadArrow': MSO_SHAPE.QUAD_ARROW,
    'mxgraph.arrows2.uTurnArrow': MSO_SHAPE.U_TURN_ARROW,
    'mxgraph.arrows2.stripedArrow': MSO_SHAPE.STRIPED_RIGHT_ARROW,
    'mxgraph.arrows2.notchedArrow': MSO_SHAPE.NOTCHED_RIGHT_ARROW,

    # Arrow callouts
    'leftArrowCallout': MSO_SHAPE.LEFT_ARROW_CALLOUT,
    'rightArrowCallout': MSO_SHAPE.RIGHT_ARROW_CALLOUT,
    'upArrowCallout': MSO_SHAPE.UP_ARROW_CALLOUT,
    'downArrowCallout': MSO_SHAPE.DOWN_ARROW_CALLOUT,
    'leftRightArrowCallout': MSO_SHAPE.LEFT_RIGHT_ARROW_CALLOUT,
    'upDownArrowCallout': MSO_SHAPE.UP_DOWN_ARROW_CALLOUT,
    'quadArrowCallout': MSO_SHAPE.QUAD_ARROW_CALLOUT,

    # ── Flowchart shapes ───────────────────────────────────────────────
    'mxgraph.flowchart.process': MSO_SHAPE.FLOWCHART_PROCESS,
    'mxgraph.flowchart.decision': MSO_SHAPE.FLOWCHART_DECISION,
    'mxgraph.flowchart.data': MSO_SHAPE.FLOWCHART_DATA,
    'mxgraph.flowchart.predefined_process': MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
    'mxgraph.flowchart.stored_data': MSO_SHAPE.FLOWCHART_STORED_DATA,
    'mxgraph.flowchart.internal_storage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
    'mxgraph.flowchart.document': MSO_SHAPE.FLOWCHART_DOCUMENT,
    'mxgraph.flowchart.multi-document': MSO_SHAPE.FLOWCHART_MULTIDOCUMENT,
    'mxgraph.flowchart.terminator': MSO_SHAPE.FLOWCHART_TERMINATOR,
    'mxgraph.flowchart.preparation': MSO_SHAPE.FLOWCHART_PREPARATION,
    'mxgraph.flowchart.manual_input': MSO_SHAPE.FLOWCHART_MANUAL_INPUT,
    'mxgraph.flowchart.manual_operation': MSO_SHAPE.FLOWCHART_MANUAL_OPERATION,
    'mxgraph.flowchart.database': MSO_SHAPE.FLOWCHART_MAGNETIC_DISK,
    'mxgraph.flowchart.direct_data': MSO_SHAPE.FLOWCHART_DIRECT_ACCESS_STORAGE,
    'mxgraph.flowchart.merge': MSO_SHAPE.FLOWCHART_MERGE,
    'mxgraph.flowchart.extract': MSO_SHAPE.FLOWCHART_EXTRACT,
    'mxgraph.flowchart.or': MSO_SHAPE.FLOWCHART_OR,
    'mxgraph.flowchart.summing_junction': MSO_SHAPE.FLOWCHART_SUMMING_JUNCTION,
    'mxgraph.flowchart.sort': MSO_SHAPE.FLOWCHART_SORT,
    'mxgraph.flowchart.collate': MSO_SHAPE.FLOWCHART_COLLATE,
    'mxgraph.flowchart.display': MSO_SHAPE.FLOWCHART_DISPLAY,
    'mxgraph.flowchart.delay': MSO_SHAPE.FLOWCHART_DELAY,
    'mxgraph.flowchart.tape_data': MSO_SHAPE.FLOWCHART_SEQUENTIAL_ACCESS_STORAGE,
    'mxgraph.flowchart.on-page_reference': MSO_SHAPE.FLOWCHART_CONNECTOR,
    'mxgraph.flowchart.off-page_reference': MSO_SHAPE.FLOWCHART_OFFPAGE_CONNECTOR,
    'mxgraph.flowchart.card': MSO_SHAPE.FLOWCHART_CARD,
    'mxgraph.flowchart.paper_tape': MSO_SHAPE.FLOWCHART_PUNCHED_TAPE,
    'mxgraph.flowchart.alternate_process': MSO_SHAPE.FLOWCHART_ALTERNATE_PROCESS,

    # Draw.io built-in flowchart aliases (non-namespaced)
    'document': MSO_SHAPE.FLOWCHART_DOCUMENT,
    'internalStorage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
    'offPageConnector': MSO_SHAPE.FLOWCHART_OFFPAGE_CONNECTOR,
    'dataStorage': MSO_SHAPE.FLOWCHART_STORED_DATA,
    'or': MSO_SHAPE.FLOWCHART_OR,
    'sumEllipse': MSO_SHAPE.FLOWCHART_SUMMING_JUNCTION,
    'display': MSO_SHAPE.FLOWCHART_DISPLAY,
    'delay': MSO_SHAPE.FLOWCHART_DELAY,
    'manualInput': MSO_SHAPE.FLOWCHART_MANUAL_INPUT,
    'card': MSO_SHAPE.FLOWCHART_CARD,
    'loopLimit': MSO_SHAPE.FLOWCHART_PREPARATION,
    'collate': MSO_SHAPE.FLOWCHART_COLLATE,
    'sortShape': MSO_SHAPE.FLOWCHART_SORT,
    'extract': MSO_SHAPE.FLOWCHART_EXTRACT,
    'merge': MSO_SHAPE.FLOWCHART_MERGE,
    'tapeData': MSO_SHAPE.FLOWCHART_SEQUENTIAL_ACCESS_STORAGE,
    'tape': MSO_SHAPE.FLOWCHART_PUNCHED_TAPE,

    # UML / misc shapes that have reasonable PPTX approximations
    'folder': MSO_SHAPE.FOLDED_CORNER,
    'image': MSO_SHAPE.RECTANGLE,
    'label': MSO_SHAPE.RECTANGLE,
    'text': MSO_SHAPE.RECTANGLE,
    'line': MSO_SHAPE.RECTANGLE,

    # Explosions / Bursts
    'mxgraph.basic.bang': MSO_SHAPE.EXPLOSION1,
    'explosion': MSO_SHAPE.EXPLOSION1,
    'explosion2': MSO_SHAPE.EXPLOSION2,

    # Diagonal stripe / Half-frame
    'diagStripe': MSO_SHAPE.DIAGONAL_STRIPE,
    'halfFrame': MSO_SHAPE.HALF_FRAME,

    # Gear
    'mxgraph.basic.gear': MSO_SHAPE.GEAR_6,
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
    """Determines the best MSO_SHAPE type based on the style dictionary.

    Lookup order:
    1. Exact match on the ``shape=`` value.
    2. Partial / suffix match on the ``shape=`` value (handles library
       prefixes like ``mxgraph.basic.`` that may differ between Draw.io
       versions).
    3. Keyword match against other style keys.
    4. ``rounded=1`` or bare ``ellipse`` flags.
    5. Aspect-ratio-aware fallback: tall-narrow shapes map to a can
       (cylinder), very wide shapes map to a rounded rectangle, and
       everything else falls back to a plain rectangle.
    """
    shape_name = style_dict.get('shape', '')

    # 1. Exact match
    if shape_name and shape_name in SHAPE_MAP:
        return SHAPE_MAP[shape_name]

    # 2. Suffix / partial match – strip namespace prefixes and retry
    if shape_name:
        suffix = shape_name.rsplit('.', 1)[-1]  # e.g. "document" from "mxgraph.flowchart.document"
        if suffix in SHAPE_MAP:
            log.debug("Shape '%s' resolved via suffix '%s'", shape_name, suffix)
            return SHAPE_MAP[suffix]

    # 3. Keyword match against other style keys
    for key in style_dict:
        if key in SHAPE_MAP:
            return SHAPE_MAP[key]

    # 4. Style flags
    if style_dict.get('rounded') == '1':
        return MSO_SHAPE.ROUNDED_RECTANGLE
    if 'ellipse' in style_dict:
        return MSO_SHAPE.OVAL

    # 5. Aspect-ratio-aware fallback
    if shape_name:
        log.info("Unmapped Draw.io shape: '%s' – using bounding-box fallback", shape_name)
    return _fallback_from_aspect(style_dict)

def _fallback_from_aspect(style_dict):
    """Choose a fallback shape based on the bounding-box aspect ratio.

    This provides a better visual approximation than always returning a
    plain rectangle.  The ``style_dict`` is expected to come from the
    parser, but width/height aren't stored there – they live on the
    vertex dict.  To keep the interface simple the caller may optionally
    set ``_width`` and ``_height`` keys before calling ``get_shape_type``.
    """
    try:
        w = float(style_dict.get('_width', 0))
        h = float(style_dict.get('_height', 0))
    except (TypeError, ValueError):
        return MSO_SHAPE.RECTANGLE

    if w > 0 and h > 0:
        ratio = h / w
        if ratio > 2.0:
            return MSO_SHAPE.CAN       # tall-narrow → cylinder approximation
        if ratio < 0.4:
            return MSO_SHAPE.ROUNDED_RECTANGLE  # very wide → rounded rect
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
