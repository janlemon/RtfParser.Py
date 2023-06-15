import os
import sys
from striprtf.striprtf import rtf_to_text
from docx import Document

def convert_rtf_to_text(rtf_path, output_path):
    # Načíst RTF soubor
    with open(rtf_path, 'rb') as file:
        rtf_data = file.read()
    
    # Převést RTF na čistý text
    try:
        text = rtf_to_text(rtf_data.decode('windows-1250'),'windows-1250')
    except UnicodeDecodeError:
        text = rtf_to_text(rtf_data.decode('latin-1'))
    
    # Uložit čistý text do souboru
    output_file = os.path.join(output_path, os.path.splitext(os.path.basename(rtf_path))[0] + '.txt')
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(text)
    
    print(f"The file has been successfully converted to plain text {output_file}")

def convert_rtf_to_docx(rtf_path, output_path):
    # Načíst RTF soubor
    with open(rtf_path, 'rb') as file:
        rtf_data = file.read()
    
    # Převést RTF na čistý text
    try:
        text = rtf_to_text(rtf_data.decode('utf-8'))
    except UnicodeDecodeError:
        text = rtf_to_text(rtf_data.decode('latin-1'))
    
    # Uložit do DOCX formátu
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
    
    # Zkontrolovat, zda je cesta k RTF souboru platná
    if not os.path.isfile(rtf_path) or not rtf_path.lower().endswith('.rtf'):
        print("Invalid path to RTF file")
        sys.exit(1)
    
    # Vytvořit výstupní složku, pokud neexistuje
    os.makedirs(output_path, exist_ok=True)
    
    # Konverze RTF na čistý text
    convert_rtf_to_text(rtf_path, output_path)
    
    # Konverze RTF na DOCX
    convert_rtf_to_docx(rtf_path, output_path)
