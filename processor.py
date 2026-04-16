import re
import base64
import io
import xml.etree.ElementTree as ET
from docx2python import docx2python
from PIL import Image
from xml.dom import minidom

# Regex yang lebih akurat
RE_OPTION = re.compile(r'^[A-E][\.\)]')
RE_IMAGE = re.compile(r'----(image\d+\.\w+)----')
RE_CLEAN_HTML = re.compile(r'<[^<]+?>')

def cdata(text):
    """Membungkus teks dalam CDATA tanpa merusak tag di dalamnya."""
    return f"<![CDATA[{text}]]>"

def optimize_image(image_bytes, max_width=600):
    try:
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        if img.size[0] > max_width:
            w_percent = (max_width / float(img.size[0]))
            h_size = int((float(img.size[1]) * float(w_percent)))
            img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=70, optimize=True)
        return buffer.getvalue(), "jpg"
    except:
        return image_bytes, "png"

def process_docx_to_moodle(docx_file_bytes, filename="QUIZ"):
    category_name = filename.rsplit('.', 1)[0].strip()
    
    with docx2python(docx_file_bytes, html=True) as doc_content:
        content = doc_content.body
        images_dict = doc_content.images
        
        quiz = ET.Element('quiz')

        # 1. KATEGORI
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = f"$course$/{category_name}"

        questions = []
        current_q = None
        q_counter = 1

        for sheet in content:
            for table in sheet:
                for row in table:
                    for cell in row:
                        for paragraph in cell:
                            text = paragraph.strip()
                            if not text: continue

                            if not RE_OPTION.match(text) and not text.startswith('ANS:'):
                                if current_q: questions.append(current_q)
                                q_type = 'essay' if 'ESSAY' in text.upper() else 'multichoice'
                                current_q = {'text': text, 'options': [], 'answer': '', 'images': [], 'type': q_type}
                            elif RE_OPTION.match(text) and current_q:
                                current_q['options'].append(RE_OPTION.sub('', text).strip())
                            elif text.startswith('ANS:') and current_q:
                                current_q['answer'] = text.replace('ANS:', '').strip().upper()

                            if '----image' in text and current_q:
                                img_refs = RE_IMAGE.findall(text)
                                for ref in img_refs:
                                    if ref in images_dict:
                                        current_q['images'].append((ref, images_dict[ref]))

        if current_q: questions.append(current_q)

        # 2. PEMBUATAN STRUKTUR XML
        for q in questions:
            q_node = ET.SubElement(quiz, 'question', type=q['type'])
            
            q_id = f"q{str(q_counter).zfill(2)}"
            pure_text = RE_CLEAN_HTML.sub('', q['text'])
            clean_snippet = pure_text[:50].strip()
            
            # Perbaikan: Jangan gunakan CDATA di dalam ET.SubElement karena akan di-escape
            # Kita akan menggantinya secara manual di akhir
            name_node = ET.SubElement(q_node, 'name')
            ET.SubElement(name_node, 'text').text = f"__CDATA_START__{category_name} {q_id} {clean_snippet}__CDATA_END__"

            qtext_tag = ET.SubElement(q_node, 'questiontext', format='html')
            html_body = f"<p dir='auto'>{q['text']}</p>"
            
            for img_name, img_bytes in q['images']:
                opt_bytes, ext = optimize_image(img_bytes)
                new_img_name = f"{q_id}_{img_name.split('.')[0]}.{ext}"
                img_html = f'<br /><img src="@@PLUGINFILE@@/{new_img_name}" border="0" /><br />'
                html_body = html_body.replace(f'----{img_name}----', img_html)
                
                file_node = ET.SubElement(qtext_tag, 'file', name=new_img_name, encoding="base64")
                file_node.text = base64.b64encode(opt_bytes).decode()

            ET.SubElement(qtext_tag, 'text').text = f"__CDATA_START__{html_body}__CDATA_END__"

            ET.SubElement(q_node, 'defaultgrade').text = "1.0"
            if q['type'] == 'multichoice':
                ET.SubElement(q_node, 'penalty').text = "0.33"
                ET.SubElement(q_node, 'hidden').text = "0"
                ET.SubElement(q_node, 'single').text = "true"
                ET.SubElement(q_node, 'shuffleanswers').text = "true"
                ET.SubElement(q_node, 'answernumbering').text = "abc"

                letters = ['A', 'B', 'C', 'D', 'E']
                for i, opt in enumerate(q['options']):
                    is_correct = "100.0" if (i < len(letters) and letters[i] in q['answer']) else "0.0"
                    ans_node = ET.SubElement(q_node, 'answer', fraction=is_correct, format="html")
                    ET.SubElement(ans_node, 'text').text = f"__CDATA_START__{opt}__CDATA_END__"
            
            elif q['type'] == 'essay':
                ET.SubElement(q_node, 'penalty').text = "0.1"
                ET.SubElement(q_node, 'responseformat').text = "editor"
                ET.SubElement(q_node, 'responserequired').text = "1"
                ET.SubElement(q_node, 'responsefieldlines').text = "15"
                ET.SubElement(q_node, 'attachments').text = "0"

            q_counter += 1

        # 3. FINALISASI & CLEANING CDATA
        raw_xml = ET.tostring(quiz, encoding='utf-8')
        # Gunakan minidom untuk Pretty Print agar rapi seperti file referensi
        pretty_xml = minidom.parseString(raw_xml).toprettyxml(indent="  ")
        
        # Kembalikan penanda __CDATA__ menjadi tag CDATA asli agar tidak ter-escape
        pretty_xml = pretty_xml.replace("__CDATA_START__", "<![CDATA[").replace("__CDATA_END__", "]]>")
        
        header_comment = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "\n"
            "\n"
            "\n"
            "\n"
        )
        
        # Hapus deklarasi XML default minidom agar tidak double
        final_output = header_comment + pretty_xml.replace('<?xml version="1.0" ?>', '').strip()
        
        return final_output
