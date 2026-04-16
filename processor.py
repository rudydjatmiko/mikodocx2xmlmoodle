import re
import base64
import xml.etree.ElementTree as ET
from xml.dom import minidom
from docx2python import docx2python

def cdata(text):
    return f"<![CDATA[{text}]]>"

def process_docx_to_moodle(docx_file_bytes):
    with docx2python(docx_file_bytes, html=True) as doc_content:
        full_body = doc_content.body
        images_dict = doc_content.images
        
        quiz = ET.Element('quiz')
        
        # Header Kategori
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = "$course$/Imported_from_App"

        questions = []
        current_q = None

        # Parsing Struktur docx2python
        for table in full_body[0]:
            for row in table:
                for cell in row:
                    for paragraph in cell:
                        text = paragraph.strip()
                        if not text: continue

                        # Deteksi Soal, Pilihan, dan Kunci
                        if not re.match(r'^[A-E][\.\)]', text) and not text.startswith('ANS:'):
                            if current_q: questions.append(current_q)
                            current_q = {'text': text, 'options': [], 'answer': '', 'images': []}
                        elif re.match(r'^[A-E][\.\)]', text) and current_q:
                            current_q['options'].append(re.sub(r'^[A-E][\.\)]', '', text).strip())
                        elif text.startswith('ANS:') and current_q:
                            current_q['answer'] = text.replace('ANS:', '').strip()

                        # Deteksi Gambar
                        if '----image' in text and current_q:
                            img_filenames = re.findall(r'----(image\d+\.\w+)----', text)
                            for fname in img_filenames:
                                if fname in images_dict:
                                    current_q['images'].append((fname, images_dict[fname]))

        if current_q: questions.append(current_q)

        # Membangun XML
        for q in questions:
            q_node = ET.SubElement(quiz, 'question', type='multichoice')
            ET.SubElement(ET.SubElement(q_node, 'name'), 'text').text = cdata(q['text'][:50])
            
            qtext_node = ET.SubElement(q_node, 'questiontext', format='html')
            html_content = f"<p dir='auto'>{q['text']}</p>"
            for img_name, img_data in q['images']:
                b64_str = base64.b64encode(img_data).decode()
                ext = img_name.split('.')[-1]
                img_tag = f'<img src="data:image/{ext};base64,{b64_str}" /><br>'
                html_content = html_content.replace(f'----{img_name}----', img_tag)
            
            ET.SubElement(qtext_node, 'text').text = cdata(html_content)

            letters = ['A', 'B', 'C', 'D', 'E']
            for i, opt in enumerate(q['options']):
                score = "100" if (i < len(letters) and letters[i] in q['answer']) else "0"
                ans_node = ET.SubElement(q_node, 'answer', fraction=score)
                ET.SubElement(ans_node, 'text').text = cdata(opt)

        return minidom.parseString(ET.tostring(quiz)).toprettyxml(indent="  ")
