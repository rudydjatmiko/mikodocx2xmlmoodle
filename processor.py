import re
import base64
import io
import xml.etree.ElementTree as ET
from docx2python import docx2python
from PIL import Image

# Pre-compile Regex untuk kecepatan eksekusi
RE_OPTION = re.compile(r'^[A-E][\.\)]')
RE_IMAGE = re.compile(r'----(image\d+\.\w+)----')
RE_CLEAN_HTML = re.compile(r'<[^<]+?>')

def cdata(text):
    """Membungkus teks dalam CDATA agar aman untuk Moodle."""
    return f"<![CDATA[{text}]]>"

def optimize_image(image_bytes, max_width=600):
    """Mengecilkan resolusi gambar untuk performa XML yang ringan."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        
        # Resize hanya jika gambar terlalu besar
        if img.size[0] > max_width:
            w_percent = (max_width / float(img.size[0]))
            h_size = int((float(img.size[1]) * float(w_percent)))
            img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
        
        buffer = io.BytesIO()
        # Menggunakan JPEG untuk kompresi maksimal (kecil & cepat)
        img.save(buffer, format="JPEG", quality=70, optimize=True)
        return buffer.getvalue(), "jpg"
    except:
        return image_bytes, "png"

def process_docx_to_moodle(docx_file_bytes, filename="QUIZ"):
    category_name = filename.rsplit('.', 1)[0].strip()
    
    # docx2python sangat efektif membaca tabel & textbox
    with docx2python(docx_file_bytes, html=True) as doc_content:
        content = doc_content.body
        images_dict = doc_content.images
        
        quiz = ET.Element('quiz')

        # 1. KATEGORI (Sesuai file referensi)
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = f"$course$/{category_name}"

        questions = []
        current_q = None
        q_counter = 1

        # Pemindaian dokumen secara linear (Efektif & Cepat)
        for sheet in content:
            for table in sheet:
                for row in table:
                    for cell in row:
                        for paragraph in cell:
                            text = paragraph.strip()
                            if not text: continue

                            # Deteksi SOAL (Autonumbering Level 0)
                            if not RE_OPTION.match(text) and not text.startswith('ANS:'):
                                if current_q: questions.append(current_q)
                                
                                q_type = 'essay' if 'ESSAY' in text.upper() else 'multichoice'
                                current_q = {
                                    'text': text,
                                    'options': [],
                                    'answer': '',
                                    'images': [],
                                    'type': q_type
                                }

                            # Deteksi PILIHAN (Autonumbering Level 1)
                            elif RE_OPTION.match(text) and current_q:
                                clean_opt = RE_OPTION.sub('', text).strip()
                                current_q['options'].append(clean_opt)

                            # Deteksi KUNCI
                            elif text.startswith('ANS:') and current_q:
                                current_q['answer'] = text.replace('ANS:', '').strip().upper()

                            # Deteksi Gambar
                            if '----image' in text and current_q:
                                img_refs = RE_IMAGE.findall(text)
                                for ref in img_refs:
                                    if ref in images_dict:
                                        current_q['images'].append((ref, images_dict[ref]))

        if current_q: questions.append(current_q)

        # 2. PEMBUATAN XML (Optimasi Kecepatan)
        for q in questions:
            q_node = ET.SubElement(quiz, 'question', type=q['type'])
            
            # Format Nama Soal: [FILE] q01 [SNIPPET]
            q_id = f"q{str(q_counter).zfill(2)}"
            clean_snippet = RE_CLEAN_HTML.sub('', q['text'])[:50]
            ET.SubElement(ET.SubElement(q_node, 'name'), 'text').text = cdata(f"{category_name} {q_id} {clean_snippet}")

            qtext_tag = ET.SubElement(q_node, 'questiontext', format='html')
            
            # HTML Body dengan support Arab (dir=auto)
            html_body = f"<p dir='auto'>{q['text']}</p>"
            
            for img_name, img_bytes in q['images']:
                opt_bytes, ext = optimize_image(img_bytes)
                new_img_name = f"{q_id}_{img_name.split('.')[0]}.{ext}"
                
                # Masukkan Tag Image standar Moodle
                img_html = f'<br /><img src="@@PLUGINFILE@@/{new_img_name}" border="0" /><br />'
                html_body = html_body.replace(f'----{img_name}----', img_html)
                
                # Masukkan file binary Base64
                file_node = ET.SubElement(qtext_tag, 'file', name=new_img_name, encoding="base64")
                file_node.text = base64.b64encode(opt_bytes).decode()

            ET.SubElement(qtext_tag, 'text').text = cdata(html_body)

            # Metadata Standar (Penalty, Grade, dsb)
            ET.SubElement(q_node, 'defaultgrade').text = "1.0"
            
            if q['type'] == 'multichoice':
                ET.SubElement(q_node, 'penalty').text = "0.33"
                ET.SubElement(q_node, 'hidden').text = "0"
                ET.SubElement(q_node, 'single').text = "true"
                ET.SubElement(q_node, 'shuffleanswers').text = "true"
                ET.SubElement(q_node, 'answernumbering').text = "abc"

                letters = ['A', 'B', 'C', 'D', 'E']
                for i, opt in enumerate(q['options']):
                    # Pecah jawaban jika ada lebih dari satu (misal: A and C)
                    is_correct = "100.0" if (i < len(letters) and letters[i] in q['answer']) else "0.0"
                    ans_node = ET.SubElement(q_node, 'answer', fraction=is_correct, format="html")
                    ET.SubElement(ans_node, 'text').text = cdata(opt)
            
            elif q['type'] == 'essay':
                ET.SubElement(q_node, 'penalty').text = "0.1"
                ET.SubElement(q_node, 'responseformat').text = "editor"
                ET.SubElement(q_node, 'responserequired').text = "1"
                ET.SubElement(q_node, 'responsefieldlines').text = "15"
                ET.SubElement(q_node, 'attachments').text = "0"

            q_counter += 1

        # 3. FINALISASI (Tanpa Pretty Print untuk Kecepatan Maksimal)
        xml_output = ET.tostring(quiz, encoding='utf-8').decode('utf-8')
        
        header_metadata = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "\n"
            "\n"
            "\n"
            "\n"
        )
        
        return header_metadata + xml_output
