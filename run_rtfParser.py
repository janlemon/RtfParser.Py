import subprocess
import sys
import os
from rtfParser import convert_rtf_to_text

def install_dependencies():
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

def main(rtf_path, output_path):
    install_dependencies()
    convert_rtf_to_text(rtf_path, output_path)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Missing input parameters. Expected path to RTF file and output directory")
        sys.exit(1)
    
    rtf_path = sys.argv[1]
    output_path = sys.argv[2]

    main(rtf_path, output_path)
