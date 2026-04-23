import xml.etree.ElementTree as ET
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import base64

def process_docx_to_xml(docx_file):
    # Memuat dokumen dari stream file
    doc = Document(docx_file)
    root = ET.Element("DocumentContent")

    # 1. Memproses Paragraf (Teks & Autonumbering)
    text_section = ET.SubElement(root, "TextContent")
    for para in doc.paragraphs:
        p_element = ET.SubElement(text_section, "Paragraph")
        
        # Deteksi sederhana untuk penomoran (List String)
        if para.style.name.startswith('List'):
            p_element.set("type", "list_item")
        
        p_element.text = para.text

    # 2. Memproses Tabel
    table_section = ET.SubElement(root, "Tables")
    for table in doc.tables:
        t_element = ET.SubElement(table_section, "Table")
        for row in table.rows:
            r_element = ET.SubElement(t_element, "Row")
            for cell in row.cells:
                c_element = ET.SubElement(r_element, "Cell")
                c_element.text = cell.text

    # 3. Memproses Gambar (Dikonversi ke Base64 agar bisa masuk XML)
    image_section = ET.SubElement(root, "Images")
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_element = ET.SubElement(image_section, "Image")
            img_element.set("name", rel.target_ref.split('/')[-1])
            
            # Encode biner gambar ke string base64
            img_data = rel.target_part.blob
            img_base64 = base64.b64encode(img_data).decode('utf-8')
            img_element.text = img_base64

    # Convert ke string XML yang rapi
    return ET.tostring(root, encoding='utf-8', method='xml')
