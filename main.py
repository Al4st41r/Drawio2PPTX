import argparse
import sys
from converter import convert

def main():
    parser = argparse.ArgumentParser(description="Convert Draw.io XML to PowerPoint")
    parser.add_argument("input_file", help="Path to input .drawio or .xml file")
    parser.add_argument("output_file", help="Path to output .pptx file")
    args = parser.parse_args()

    try:
        convert(args.input_file, args.output_file)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
