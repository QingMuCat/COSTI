import zipfile
import xml.etree.ElementTree as ET
import os

def extract_text_from_docx(docx_path):
    """Extract text from a .docx file"""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            # Read the main document XML
            document_xml = zip_ref.read('word/document.xml')
            
            # Parse the XML
            root = ET.fromstring(document_xml)
            
            # Define namespaces
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }
            
            # Extract all paragraph text
            paragraphs = root.findall('.//w:p', ns)
            text_content = []
            
            for para in paragraphs:
                # Get all text elements in the paragraph
                text_elements = para.findall('.//w:t', ns)
                para_text = ''.join([elem.text or '' for elem in text_elements])
                if para_text.strip():
                    text_content.append(para_text.strip())
            
            return '\n'.join(text_content)
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == '__main__':
    docx_path = r'Q:\HuaweiMoveData\Users\12425\Desktop\需求文档.docx'
    text = extract_text_from_docx(docx_path)
    
    # Save to file
    output_path = r'C:\Users\12425\lobsterai\project\需求文档内容.txt'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(text)
    
    print(f"Content extracted to {output_path}")
    print("\n--- Document Content ---\n")
    print(text)
