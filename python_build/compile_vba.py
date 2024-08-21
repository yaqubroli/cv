from vbaProjectCompiler.vbaProject import VbaProject
from vbaProjectCompiler.ole_file import OleFile

def compile_vba(input_dir, docm_file):
    # Initialize the VbaProject object
    vba_project = VbaProject()

    # Load the VBA project from the extracted directory
    vba_project.addDirectory(input_dir)

    # Create the OLE file with the VBA project
    ole_file = OleFile(vba_project)
    ole_file.writeFile(docm_file)

    print(f"VBA project compiled and saved to {docm_file}")

if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python compile_vba.py <input_directory> <output_docm_file>")
        sys.exit(1)

    input_dir = sys.argv[1]
    docm_file = sys.argv[2]

    compile_vba(input_dir, docm_file)
