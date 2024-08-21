from oletools.olevba import VBA_Parser
import os

def extract_vba(docm_file, output_dir):
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Parse the .docm file to extract VBA code
    vba_parser = VBA_Parser(docm_file)
    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            # Save each extracted VBA file to the output directory
            output_file_path = os.path.join(output_dir, vba_filename)
            with open(output_file_path, 'w') as vba_file:
                vba_file.write(vba_code)
                print(f"Extracted {vba_filename} to {output_file_path}")
    else:
        print("No VBA macros found in the document.")

    vba_parser.close()

if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python extract_vba.py <input_docm_file> <output_directory>")
        sys.exit(1)

    docm_file = sys.argv[1]
    output_dir = sys.argv[2]

    extract_vba(docm_file, output_dir)
