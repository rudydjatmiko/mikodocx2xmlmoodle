import re
import base64
import io
import xml.etree.ElementTree as ET
from xml.dom import minidom
from docx2python import docx2python
from PIL import Image

def cdata(text):
    """Membungkus teks dalam CDATA agar aman bagi Moodle (untuk Arab & Rumus)."""
    return f"<![CDATA[{text}]]>"

def optimize_image(image_bytes, max_width=600):
    """Mengecilkan resolusi gambar agar ukuran file XML tidak terlalu besar."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        # Konversi ke RGB jika formatnya RGBA (agar bisa disave ke JPEG)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        
        # Hitung proporsi
        w_percent = (max_width / float(img.size[0]))
        if w_percent < 1.0:
            h_size = int((float(img.size[1]) * float(w_percent)))
            img = img.resize((max_width, h_size), Image.Resampling.LANCZOS)
        
        # Simpan ke format JPEG dengan kualitas 70%
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=70, optimize=True)
        return buffer.getvalue(), "jpg"
    except Exception:
        # Jika gagal optimasi, return data asli
        return image_bytes, "png"

def process_docx_to_moodle(docx_file_bytes, filename="QUIZ"):
    # Membersihkan nama kategori dari nama file
    category_name = filename.replace('.docx', '').replace('.doc', '').strip()
    
    # Ekstraksi konten dengan docx2python (html=True untuk menjaga format Bold/Italic/Table)
    with docx2python(docx_file_bytes, html=True) as doc_content:
        content = doc_content.body
        images_dict = doc_content.images
        
        quiz = ET.Element('quiz')

        # 1. HEADER KATEGORI
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = f"$course$/{category_name}"

        questions = []
        current_q = None
        q_counter = 1

        # Pemindaian konten dokumen
        for sheet in content:
            for table in sheet:
                for row in table:
                    for cell in row:
                        for paragraph in cell:
                            text = paragraph.strip()
                            if not text: continue

                            # LOGIKA DETEKSI SOAL
                            # Bukan pilihan (A-E) dan bukan Kunci Jawaban (ANS:)
                            if not re.match(r'^[A-E][\.\)]', text) and not text.startswith('ANS:'):
                                if current_q: questions.append(current_q)
                                
                                q_type = 'essay' if 'ESSAY' in text.upper() else 'multichoice'
                                current_q = {
                                    'text': text,
                                    'options': [],
                                    'answer': '',
                                    'images': [],
                                    'type': q_type
                                }

                            # LOGIKA DETEKSI PILIHAN
                            elif re.match(r'^[A-E][\.\)]', text) and current_q:
                                clean_opt = re.sub(r'^[A-E][\.\)]', '', text).strip()
                                current_q['options'].append(clean_opt)

                            # LOGIKA DETEKSI KUNCI
                            elif text.startswith('ANS:') and current_q:
                                current_q['answer'] = text.replace('ANS:', '').strip().upper()

                            # LOGIKA DETEKSI GAMBAR
                            if '----image' in text and current_q:
                                img_refs = re.findall(r'----(image\d+\.\w+)----', text)
                                for ref in img_refs:
                                    if ref in images_dict:
                                        current_q['images'].append((ref, images_dict[ref]))

        if current_q: questions.append(current_q)

        # 2. KONSTRUKSI TIAP SOAL KE XML
        for q in questions:
            q_node = ET.SubElement(quiz, 'question', type=q['type'])
            
            # Name: Sesuai file referensi [PREFIX] q01 [SNIPPET]
            q_id = f"q{str(q_counter).zfill(2)}"
            # Hapus tag HTML dari snippet nama agar XML valid
            clean_snippet = re.sub('<[^<]+?>', '', q['text'])[:50]
            q_name_text = f"{category_name} {q_id} {clean_snippet}"
            ET.SubElement(ET.SubElement(q_node, 'name'), 'text').text = cdata(q_name_text)

            # Question Text
            qtext_tag = ET.SubElement(q_node, 'questiontext', format='html')
            
            # Proses Gambar & Teks (Support Arab dir=auto)
            html_body = f"<p dir='auto'>{q['text']}</p>"
            
            for img_name, img_bytes in q['images']:
                # Kompres gambar sebelum diencode
                opt_bytes, ext = optimize_image(img_bytes)
                new_img_name = f"{q_id}_{img_name.split('.')[0]}.{ext}"
                
                # Ganti placeholder di text dengan tag img
                img_html = f'<br /><img src="@@PLUGINFILE@@/{new_img_name}" border="0" /><br />'
                html_body = html_body.replace(f'----{img_name}----', img_html)
                
                # Masukkan file binary (Base64)
                file_node = ET.SubElement(qtext_tag, 'file', name=new_img_name, encoding="base64")
                file_node.text = base64.b64encode(opt_bytes).decode()

            ET.SubElement(qtext_tag, 'text').text = cdata(html_body)

            # Metadata Tambahan
            ET.SubElement(q_node, 'defaultgrade').text = "1.0"
            
            if q['type'] == 'multichoice':
                ET.SubElement(q_node, 'penalty').text = "0.33"
                ET.SubElement(q_node, 'hidden').text = "0"
                ET.SubElement(q_node, 'single').text = "true"
                ET.SubElement(q_node, 'shuffleanswers').text = "true"
                ET.SubElement(q_node, 'answernumbering').text = "abc"

                # Pilihan Jawaban
                letters = ['A', 'B', 'C', 'D', 'E']
                for i, opt in enumerate(q['options']):
                    # Mendukung multiple answer jika kunci mengandung lebih dari 1 huruf
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

        # 3. FINALISASI & HEADER GARY BLACKBURN
        xml_string = ET.tostring(quiz, encoding='utf-8')
        # Gunakan minidom hanya untuk merapikan spasi (indentasi)
        pretty_xml = minidom.parseString(xml_string).toprettyxml(indent="  ")
        
        header_metadata = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "\n"
            "\n"
            "\n"
            "\n"
        )
        
        # Gabungkan dan hapus deklarasi XML ganda dari minidom
        final_output = header_metadata + pretty_xml.replace('<?xml version="1.0" ?>', '').strip()
        
        return final_output
