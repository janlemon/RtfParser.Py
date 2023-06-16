import os
import sys
from striprtf.striprtf import rtf_to_text
from docx import Document

def convert_rtf_to_text(rtf_path, output_path):
    # Load RTF File
    with open(rtf_path, 'rb') as file:
        rtf_data = file.read()
    
    # Convert RTF File to plain text
    # Using UnicodeDecode for Czech Republic letters
    try:
        text = rtf_to_text(rtf_data.decode('windows-1250'),'windows-1250')
    except UnicodeDecodeError:
        text = rtf_to_text(rtf_data.decode('latin-1'))
    
    # Save plain text to file
    output_file = os.path.join(output_path, os.path.splitext(os.path.basename(rtf_path))[0] + '.txt')
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(text)
    
    print(f"The file has been successfully converted to plain text {output_file}")

#Can be used for docx convertor, not used in main function
def convert_rtf_to_docx(rtf_path, output_path):
    # Load RTF File
    with open(rtf_path, 'rb') as file:
        rtf_data = file.read()
    
    # Convert RTF File to plain text
    # Using UnicodeDecode for Czech Republic letters
    try:
        text = rtf_to_text(rtf_data.decode('windows-1250'),'windows-1250')
    except UnicodeDecodeError:
        text = rtf_to_text(rtf_data.decode('latin-1'))
    
    # Save docx to file
    output_file = os.path.join(output_path, os.path.splitext(os.path.basename(rtf_path))[0] + '.docx')
    doc = Document()
    doc.add_paragraph(text)
    doc.save(output_file)
    
    print(f"The file has been successfully converted to txt {output_file}")

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Missing input parameters. Expected path to RTF file and output directory")
        sys.exit(1)
    
    rtf_path = sys.argv[1]
    output_path = sys.argv[2]
    
    # Check if file path is true
    if not os.path.isfile(rtf_path) or not rtf_path.lower().endswith('.rtf'):
        print("Invalid path to RTF file")
        sys.exit(1)
    
    # Create a folder, if folder doesn't exist
    os.makedirs(output_path, exist_ok=True)
    
    # Konverze RTF na čistý text
    convert_rtf_to_text(rtf_path, output_path)
    
    # RTF to Docx conversion
    convert_rtf_to_docx(rtf_path, output_path)
