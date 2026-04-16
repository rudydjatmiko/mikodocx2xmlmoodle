import re
import base64
import xml.etree.ElementTree as ET
from xml.dom import minidom
from docx2python import docx2python

def cdata(text):
    """Membungkus teks dalam CDATA agar aman untuk Moodle."""
    return f"<![CDATA[{text}]]>"

def process_docx_to_moodle(docx_file_bytes, filename="QUIZ"):
    # Hapus ekstensi dari nama file untuk kategori
    category_name = filename.replace('.docx', '').replace('.doc', '')
    
    with docx2python(docx_file_bytes, html=True) as doc_content:
        # doc_content.body[0] biasanya berisi tabel utama dokumen
        content = doc_content.body
        images_dict = doc_content.images
        
        quiz = ET.Element('quiz')

        # 1. TAMBAHKAN KATEGORI (Sesuai file referensi)
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = f"$course$/{category_name}"

        questions = []
        current_q = None
        q_counter = 1

        # Flatten structure dari docx2python (Sheets > Tables > Rows > Cells > Paragraphs)
        for sheet in content:
            for table in sheet:
                for row in table:
                    for cell in row:
                        for paragraph in cell:
                            text = paragraph.strip()
                            if not text: continue

                            # DETEKSI SOAL (Bukan pilihan A-E dan bukan ANS:)
                            # Menangani autonumbering atau teks biasa sebagai soal
                            if not re.match(r'^[A-E][\.\)]', text) and not text.startswith('ANS:'):
                                # Jika ada kata ESSAY, tandai sebagai tipe essay
                                q_type = 'essay' if 'ESSAY' in text.upper() else 'multichoice'
                                
                                if current_q: questions.append(current_q)
                                
                                current_q = {
                                    'text': text,
                                    'options': [],
                                    'answer': '',
                                    'images': [],
                                    'type': q_type
                                }

                            # DETEKSI PILIHAN JAWABAN (A. B. C. D.)
                            elif re.match(r'^[A-E][\.\)]', text) and current_q:
                                clean_opt = re.sub(r'^[A-E][\.\)]', '', text).strip()
                                current_q['options'].append(clean_opt)

                            # DETEKSI KUNCI JAWABAN (ANS: B)
                            elif text.startswith('ANS:') and current_q:
                                # Mengambil huruf setelah 'ANS:'
                                current_q['answer'] = text.replace('ANS:', '').strip()

                            # DETEKSI PLACEHOLDER GAMBAR
                            if '----image' in text and current_q:
                                img_filenames = re.findall(r'----(image\d+\.\w+)----', text)
                                for fname in img_filenames:
                                    if fname in images_dict:
                                        current_q['images'].append((fname, images_dict[fname]))

        if current_q: questions.append(current_q)

        # 2. KONSTRUKSI XML PER SOAL
        for q in questions:
            q_node = ET.SubElement(quiz, 'question', type=q['type'])
            
            # Name: Format [CATEGORY] q01 [TEXT SNIPPET]
            q_id = f"q{str(q_counter).zfill(2)}"
            q_name_text = f"{category_name} {q_id} {q['text'][:50]}..."
            ET.SubElement(ET.SubElement(q_node, 'name'), 'text').text = cdata(q_name_text)

            # Question Text
            qtext_node = ET.SubElement(q_node, 'questiontext', format='html')
            
            # Format HTML Soal (Mendukung Font Arab/Matematika via CDATA)
            html_content = f"<p dir='auto'>{q['text']}</p>"
            
            # Masukkan Gambar jika ada
            for img_name, img_data in q['images']:
                # Moodle Path @@PLUGINFILE@@
                img_tag = f'<br /><img src="@@PLUGINFILE@@/{img_name}" border="0" /><br />'
                html_content = html_content.replace(f'----{img_name}----', img_tag)
                
                # Masukkan data Base64 tepat di bawah questiontext
                file_node = ET.SubElement(qtext_node, 'file', name=img_name, encoding="base64")
                file_node.text = base64.b64encode(img_data).decode()

            ET.SubElement(qtext_node, 'text').text = cdata(html_content)

            # Metadata Standar (Penalty, Grade, dsb)
            ET.SubElement(q_node, 'defaultgrade').text = "1.0"
            
            if q['type'] == 'multichoice':
                ET.SubElement(q_node, 'penalty').text = "0.33"
                ET.SubElement(q_node, 'hidden').text = "0"
                ET.SubElement(q_node, 'single').text = "true"
                ET.SubElement(q_node, 'shuffleanswers').text = "true"
                ET.SubElement(q_node, 'answernumbering').text = "abc"

                # Tambahkan Pilihan Jawaban
                letters = ['A', 'B', 'C', 'D', 'E']
                for i, opt_text in enumerate(q['options']):
                    # Bandingkan huruf (A, B, C...) dengan isi current_q['answer']
                    is_correct = "100.0" if (i < len(letters) and letters[i] in q['answer']) else "0.0"
                    
                    ans_node = ET.SubElement(q_node, 'answer', fraction=is_correct, format="html")
                    ET.SubElement(ans_node, 'text').text = cdata(opt_text)
            
            elif q['type'] == 'essay':
                ET.SubElement(q_node, 'penalty').text = "0.1"
                ET.SubElement(q_node, 'responseformat').text = "editor"
                ET.SubElement(q_node, 'responserequired').text = "1"
                ET.SubElement(q_node, 'responsefieldlines').text = "15"
                ET.SubElement(q_node, 'attachments').text = "0"

            q_counter += 1

        # 3. FORMATING AKHIR & HEADER METADATA
        raw_xml = ET.tostring(quiz, encoding='utf-8')
        reparsed = minidom.parseString(raw_xml)
        pretty_xml = reparsed.toprettyxml(indent="  ")

        # Tambahkan Header Komentar Gary Blackburn sesuai request
        header_comment = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "\n"
            "\n"
            "\n"
            "\n"
        )
        
        # Gabungkan (Hapus deklarasi XML default dari minidom agar tidak double)
        final_xml = header_comment + pretty_xml.replace('<?xml version="1.0" ?>', '').strip()
        
        return final_xml
