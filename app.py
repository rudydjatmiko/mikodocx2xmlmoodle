import streamlit as st
import re
import base64
import io
from docx2python import docx2python
import xml.etree.ElementTree as ET
from xml.dom import minidom

# Fungsi untuk membungkus konten dalam CDATA agar karakter khusus (Arabic/Math) tidak error
def cdata(text):
    return f"<![CDATA[{text}]]>"

def process_docx_to_moodle(docx_file_bytes):
    # Menggunakan docx2python untuk ekstraksi mendalam
    # html=True mempertahankan format dasar seperti bold, italic, dan tabel
    with docx2python(docx_file_bytes, html=True) as doc_content:
        # doc_content.body adalah list bertingkat: [sheet][table][row][cell][paragraph]
        full_body = doc_content.body
        images_dict = doc_content.images  # Dictionary berisi {nama_file: binary_data}
        
        quiz = ET.Element('quiz')
        
        # Header Kategori
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = "$course$/Imported_from_App"

        questions = []
        current_q = None

        # Iterasi melalui konten (docx2python meratakan struktur menjadi list)
        # Kita ambil body utama (index 0 jika tidak ada header/footer khusus)
        for table in full_body[0]:
            for row in table:
                for cell in row:
                    for paragraph in cell:
                        text = paragraph.strip()
                        if not text:
                            continue

                        # 1. DETEKSI SOAL (Level 0 / Teks Utama)
                        # Logika: Jika tidak diawali pilihan A-E dan bukan kunci jawaban
                        if not re.match(r'^[A-E][\.\)]', text) and not text.startswith('ANS:'):
                            if current_q:
                                questions.append(current_q)
                            current_q = {
                                'text': text,
                                'options': [],
                                'answer': '',
                                'images': []
                            }

                        # 2. DETEKSI PILIHAN JAWABAN (Level 1 / A. B. C.)
                        elif re.match(r'^[A-E][\.\)]', text) and current_q:
                            clean_opt = re.sub(r'^[A-E][\.\)]', '', text).strip()
                            current_q['options'].append(clean_opt)

                        # 3. DETEKSI KUNCI JAWABAN (ANS: C)
                        elif text.startswith('ANS:') and current_q:
                            current_q['answer'] = text.replace('ANS:', '').strip()

                        # 4. DETEKSI GAMBAR dalam paragraf
                        if '----image' in text and current_q:
                            img_filenames = re.findall(r'----(image\d+\.\w+)----', text)
                            for fname in img_filenames:
                                if fname in images_dict:
                                    current_q['images'].append((fname, images_dict[fname]))

        if current_q:
            questions.append(current_q)

        # Membangun XML Moodle
        for q in questions:
            q_node = ET.SubElement(quiz, 'question', type='multichoice')
            
            # Judul Soal (diambil dari potongan teks)
            name = ET.SubElement(q_node, 'name')
            ET.SubElement(name, 'text').text = cdata(q['text'][:50])

            # Isi Soal
            qtext_node = ET.SubElement(q_node, 'questiontext', format='html')
            
            # Proses gambar menjadi tag HTML <img> dengan Base64
            html_content = f"<p dir='auto'>{q['text']}</p>"
            for img_name, img_data in q['images']:
                b64_str = base64.b64encode(img_data).decode()
                ext = img_name.split('.')[-1]
                img_tag = f'<img src="data:image/{ext};base64,{b64_str}" /><br>'
                # Ganti placeholder image dari docx2python menjadi tag img nyata
                html_content = html_content.replace(f'----{img_name}----', img_tag)

            ET.SubElement(qtext_node, 'text').text = cdata(html_content)

            # Pilihan Jawaban
            letters = ['A', 'B', 'C', 'D', 'E']
            for i, opt in enumerate(q['options']):
                score = "100" if (i < len(letters) and letters[i] in q['answer']) else "0"
                ans_node = ET.SubElement(q_node, 'answer', fraction=score)
                ET.SubElement(ans_node, 'text').text = cdata(opt)

        # Format XML agar rapi
        xml_str = minidom.parseString(ET.tostring(quiz)).toprettyxml(indent="  ")
        return xml_str

# --- TAMPILAN STREAMLIT ---
st.set_page_config(page_title="Converter Soal Moodle", page_icon="📝")

st.title("📝 Docx to Moodle XML Converter")
st.markdown("""
Aplikasi ini mengonversi file Word ke format Moodle XML.
**Fitur:** Gambar, Tabel, Font Arab (Auto-dir), dan Rumus.
""")

file_upload = st.file_uploader("Unggah file .docx", type=["docx"])

if file_upload:
    if st.button("Konversi Sekarang"):
        try:
            # Baca file sebagai bytes untuk docx2python
            file_bytes = file_upload.read()
            xml_result = process_docx_to_moodle(io.BytesIO(file_bytes))
            
            st.success("Konversi Berhasil!")
            st.download_button(
                label="📥 Unduh File XML",
                data=xml_result,
                file_name=file_upload.name.replace(".docx", ".xml"),
                mime="text/xml"
            )
            
            with st.expander("Lihat Hasil XML"):
                st.code(xml_result, language="xml")
                
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
