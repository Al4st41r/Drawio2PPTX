from .parser import DrawioParser
from .engine import PptxGenerator

__version__ = "1.0.8"

def convert(input_file, output_file):
    print(f"Parsing {input_file}...")
    parser = DrawioParser(input_file)
    pages = parser.parse()
    
    print(f"Found {len(pages)} pages.")
    
    print(f"Generating {output_file}...")
    generator = PptxGenerator(output_file)
    
    for page in pages:
        name = page['name']
        vertices, edges = page['data']
        print(f"  Processing page '{name}': {len(vertices)} shapes, {len(edges)} connections")
        
        generator.create_slide()
        # TODO: Set slide title if we add title support
        generator.add_vertices(vertices)
        generator.add_edges(edges)
        
    generator.save()
    print("Done.")
