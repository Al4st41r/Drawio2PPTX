import argparse
import sys
from converter import convert, __version__

def main():
    parser = argparse.ArgumentParser(description="Convert Draw.io XML to PowerPoint")
    parser.add_argument("input_file", nargs='?', help="Path to input .drawio or .xml file")
    parser.add_argument("output_file", nargs='?', help="Path to output .pptx file")
    parser.add_argument("--version", action="version", version=f"Drawio2PPTX {__version__}")
    args = parser.parse_args()

    if not args.input_file or not args.output_file:
        parser.print_help()
        sys.exit(1)

    try:
        convert(args.input_file, args.output_file)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
