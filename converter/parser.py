import xml.etree.ElementTree as ET
from .utils import parse_style_string

class DrawioParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.root = None
        self.graph_model = None
        
    def parse(self):
        try:
            tree = ET.parse(self.file_path)
        except Exception as e:
            raise ValueError(f"Error parsing XML: {e}")

        root = tree.getroot()
        
        # Draw.io structure: mxfile -> diagram -> mxGraphModel -> root -> mxCell
        diagram = root.find('diagram')
        if diagram is None:
            self.graph_model = root.find('mxGraphModel') or root.find('.//mxGraphModel')
        else:
            self.graph_model = diagram.find('mxGraphModel')
            
        if self.graph_model is None:
            raise ValueError("Could not find mxGraphModel")
            
        return self._extract_elements()
        
    def _extract_elements(self):
        root_cell = self.graph_model.find('root')
        cells = root_cell.findall('mxCell')
        
        vertices = []
        edges = []
        
        for cell in cells:
            attrib = cell.attrib
            cell_data = {
                'id': attrib.get('id'),
                'value': attrib.get('value', ''),
                'style_str': attrib.get('style', ''),
                'style': parse_style_string(attrib.get('style', '')),
                'geometry': cell.find('mxGeometry'),
                'vertex': attrib.get('vertex') == '1',
                'edge': attrib.get('edge') == '1',
                'source': attrib.get('source'),
                'target': attrib.get('target')
            }
            
            if cell_data['vertex']:
                # Parse geometry for vertices
                geo = cell_data['geometry']
                if geo is not None:
                    cell_data['x'] = float(geo.get('x', 0))
                    cell_data['y'] = float(geo.get('y', 0))
                    cell_data['width'] = float(geo.get('width', 0))
                    cell_data['height'] = float(geo.get('height', 0))
                    vertices.append(cell_data)
            
            elif cell_data['edge']:
                edges.append(cell_data)
                
        return vertices, edges
