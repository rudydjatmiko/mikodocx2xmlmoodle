import streamlit as st
import io
from processor import process_docx_to_moodle

st.set_page_config(page_title="Converter Soal Moodle", page_icon="📝")

st.title("📝 Docx to Moodle XML Converter")
st.write("Sistem Konversi otomatis untuk Soal, Gambar, dan Arab.")

file_upload = st.file_uploader("Unggah file .docx", type=["docx"])

if file_upload:
    if st.button("🚀 Jalankan Konversi"):
        try:
            # Panggil fungsi dari file processor.py
            xml_result = process_docx_to_moodle(io.BytesIO(file_upload.read()))
            
            st.success("Konversi Berhasil!")
            st.download_button(
                label="📥 Unduh File XML",
                data=xml_result,
                file_name=file_upload.name.replace(".docx", ".xml"),
                mime="text/xml"
            )
            
            with st.expander("Preview Hasil XML"):
                st.code(xml_result, language="xml")
                
        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {e}")
