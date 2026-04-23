import streamlit as st
from processor import process_docx_to_xml

st.set_page_config(page_title="Docx to XML Converter", layout="wide")

st.title("📄 Docx to XML Processor")
st.write("Ekstrak teks, tabel, dan gambar dari file Word ke format XML.")

uploaded_file = st.file_uploader("Pilih file .docx", type="docx")

if uploaded_file is not None:
    try:
        with st.spinner('Memproses dokumen...'):
            # Panggil fungsi dari processor.py
            xml_data = process_docx_to_xml(uploaded_file)
            
            # Tampilkan Preview XML
            st.subheader("Preview XML")
            st.code(xml_data.decode('utf-8'), language='xml')
            
            # Tombol Download
            st.download_button(
                label="Download XML File",
                data=xml_data,
                file_name=f"{uploaded_file.name.split('.')[0]}.xml",
                mime="application/xml"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
