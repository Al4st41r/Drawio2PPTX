from .parser import DrawioParser
from .engine import PptxGenerator

def convert(input_file, output_file):
    print(f"Parsing {input_file}...")
    parser = DrawioParser(input_file)
    vertices, edges = parser.parse()
    
    print(f"Found {len(vertices)} shapes and {len(edges)} connections.")
    
    print(f"Generating {output_file}...")
    generator = PptxGenerator(output_file)
    generator.add_vertices(vertices)
    generator.add_edges(edges)
    generator.save()
    print("Done.")
