import xml.etree.ElementTree as ET
from .utils import parse_style_string

class DrawioParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.root = None
        
    def parse(self):
        try:
            tree = ET.parse(self.file_path)
        except Exception as e:
            raise ValueError(f"Error parsing XML: {e}")

        self.root = tree.getroot()
        
        pages = []
        
        # Draw.io structure: mxfile -> diagram
        diagrams = self.root.findall('diagram')
        
        if not diagrams:
            # Fallback for single page or raw model without diagram tag
            graph_model = self.root.find('mxGraphModel') or self.root.find('.//mxGraphModel')
            if graph_model is not None:
                pages.append({
                    'name': 'Page-1',
                    'data': self._extract_elements(graph_model)
                })
        else:
            for diagram in diagrams:
                name = diagram.get('name', 'Page')
                graph_model = diagram.find('mxGraphModel')
                # TODO: Add decompression logic here if graph_model is missing but text content exists
                if graph_model is not None:
                    pages.append({
                        'name': name,
                        'data': self._extract_elements(graph_model)
                    })
            
        if not pages:
            raise ValueError("Could not find any mxGraphModel")
            
        return pages
        
    def _extract_elements(self, graph_model):
        root_cell = graph_model.find('root')
        if root_cell is None:
            return [], []
            
        cells = root_cell.findall('mxCell')
        
        # Map for easy lookup
        cell_map = {c.get('id'): c for c in cells}
        
        vertices = []
        edges = []
        
        def get_abs_pos(cell_id):
            """Recursively calculate absolute x, y."""
            cell = cell_map.get(cell_id)
            if cell is None:
                return 0, 0
            
            parent_id = cell.get('parent')
            px, py = 0, 0
            if parent_id and parent_id in cell_map:
                # We skip the '0' and '1' root/layer cells which usually have no geometry
                parent_cell = cell_map[parent_id]
                if parent_cell.get('vertex') == '1':
                    px, py = get_abs_pos(parent_id)
            
            geo = cell.find('mxGeometry')
            if geo is not None:
                try:
                    # For vertices, x and y are absolute or relative to parent
                    # Note: mxGeometry for edges might have x,y as label offsets
                    # We only care about vertex positioning here
                    x = float(geo.get('x', 0))
                    y = float(geo.get('y', 0))
                    return px + x, py + y
                except ValueError:
                    pass
            return px, py

        for cell in cells:
            attrib = cell.attrib
            cell_id = attrib.get('id')
            
            cell_data = {
                'id': cell_id,
                'value': attrib.get('value', ''),
                'style_str': attrib.get('style', ''),
                'style': parse_style_string(attrib.get('style', '')),
                'geometry': cell.find('mxGeometry'),
                'vertex': attrib.get('vertex') == '1',
                'edge': attrib.get('edge') == '1',
                'source': attrib.get('source'),
                'target': attrib.get('target'),
                'parent_id': attrib.get('parent')
            }
            
            if cell_data['vertex']:
                geo = cell_data['geometry']
                if geo is not None:
                    # Calculate absolute position
                    ax, ay = get_abs_pos(cell_id)
                    cell_data['x'] = ax
                    cell_data['y'] = ay
                    cell_data['width'] = float(geo.get('width', 0))
                    cell_data['height'] = float(geo.get('height', 0))
                    vertices.append(cell_data)
            
            elif cell_data['edge']:
                edges.append(cell_data)
                
        return vertices, edges
