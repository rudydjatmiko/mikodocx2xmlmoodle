import streamlit as st
import logging
import re
from pathlib import Path
from processor import process_docx_to_moodle, ConversionError
import xml.etree.ElementTree as ET

# ============ LOGGING SETUP ============
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============ CONSTANTS ============
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
PREVIEW_LINES = 150
PREVIEW_SIZE_THRESHOLD = 5 * 1024 * 1024  # 5MB

# ============ PAGE CONFIG ============
st.set_page_config(
    page_title="Converter Soal Moodle",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📝 Docx to Moodle XML Converter")
st.write("Sistem Konversi otomatis untuk Soal, Gambar, dan Format Moodle")

# ============ SIDEBAR - PANDUAN ============
with st.sidebar:
    st.header("📖 Panduan Penggunaan")
    
    with st.expander("📋 Format File DOCX", expanded=True):
        st.markdown("""
        ### Soal Pilihan Ganda (Multiple Choice)
        ```
        Berapa hasil 2 + 2?
        A. 3
        B. 4
        C. 5
        D. 6
        ANS: B
        ```
        
        ### Soal Essay
        ```
        Jelaskan konsep fotosintesis
        ESSAY
        ```
        
        ### Dengan Gambar
        ```
        Lihat gambar di bawah:
        ----image1.png----
        
        A. Opsi pertama
        B. Opsi kedua
        ANS: A
        ```
        
        ### Catatan:
        - Opsi harus dimulai A-E (maksimal 5)
        - Jawaban ditandai dengan "ANS: [huruf]"
        - Minimal 2 opsi untuk pilihan ganda
        """)
    
    with st.expander("⚙️ Pengaturan Lanjutan"):
        st.markdown("""
        **Ukuran file maksimal:** 50 MB
        
        **Format gambar:** JPG, PNG
        
        **Lebar gambar default:** 600px
        
        **Kualitas gambar:** 75%
        """)
    
    with st.divider()
    
    st.markdown("**🔗 Link Moodle:**")
    st.markdown("[Dokumentasi Moodle](https://docs.moodle.org/)")

# ============ MAIN CONTENT ============
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📤 Unggah File")
    file_upload = st.file_uploader(
        "Pilih file .docx untuk dikonversi",
        type=["docx"],
        help="Format: Microsoft Word (.docx)"
    )

with col2:
    st.subheader("💡 Info")
    st.info(
        "✅ Konversi offline (aman)\n"
        "⚡ Cepat dan akurat\n"
        "🖼️ Support gambar"
    )

# ============ FILE VALIDATION ============
if file_upload:
    # Display file info
    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📝 Nama File", file_upload.name)
    with col2:
        file_size_mb = file_upload.size / (1024 * 1024)
        st.metric("📊 Ukuran", f"{file_size_mb:.2f} MB")
    with col3:
        st.metric("🔍 Tipe", "DOCX")
    
    st.divider()
    
    # Validate file size
    if file_upload.size > MAX_FILE_SIZE:
        st.error(
            f"❌ **File terlalu besar!**\n\n"
            f"Ukuran: {file_upload.size / (1024*1024):.2f} MB\n"
            f"Maksimal: {MAX_FILE_SIZE / (1024*1024):.0f} MB"
        )
    else:
        # Validate DOCX format
        file_bytes = file_upload.read()
        
        if not file_bytes.startswith(b'PK'):
            st.error("❌ **File bukan DOCX yang valid**\n\nFile harus berformat Microsoft Word (.docx)")
        else:
            # ============ CONVERSION BUTTON ============
            if st.button("🚀 Jalankan Konversi", use_container_width=True, type="primary"):
                try:
                    with st.spinner("⏳ Sedang mengkonversi... Mohon tunggu"):
                        result = process_docx_to_moodle(file_bytes, file_upload.name)
                    
                    if not result['success']:
                        st.error("❌ Konversi gagal (tidak ada soal valid)")
                    else:
                        xml_result = result['xml']
                        stats = result['stats']
                        
                        # ============ SUCCESS MESSAGE ============
                        st.success("✅ Konversi Berhasil!")
                        
                        # ============ STATISTICS ============
                        st.subheader("📊 Statistik Konversi")
                        
                        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
                        
                        with metric_col1:
                            st.metric(
                                "📋 Soal",
                                stats['total_questions'],
                                delta="soal terkonversi"
                            )
                        
                        with metric_col2:
                            st.metric(
                                "🖼️ Gambar",
                                stats['total_images'],
                                delta="gambar terproses"
                            )
                        
                        with metric_col3:
                            output_size = len(xml_result) / 1024
                            st.metric(
                                "📦 Ukuran Output",
                                f"{output_size:.2f} KB",
                                delta="XML file"
                            )
                        
                        with metric_col4:
                            warning_count = len(stats['warnings'])
                            st.metric(
                                "⚠️ Peringatan",
                                warning_count,
                                delta="issues" if warning_count > 0 else "aman"
                            )
                        
                        # ============ DOWNLOAD BUTTON ============
                        st.divider()
                        
                        safe_filename = re.sub(
                            r'[<>:"/\\|?*]', '_',
                            Path(file_upload.name).stem
                        )[:100]
                        output_filename = f"{safe_filename}.xml"
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.download_button(
                                label="📥 Unduh File XML",
                                data=xml_result,
                                file_name=output_filename,
                                mime="text/xml",
                                use_container_width=True
                            )
                        
                        with col2:
                            # Copy XML button
                            st.button(
                                "📋 Copy ke Clipboard",
                                help="Salin XML ke clipboard",
                                use_container_width=True,
                                on_click=lambda: st.write("XML copied!") if len(xml_result) < 100000 else None
                            )
                        
                        # ============ WARNINGS & ERRORS ============
                        if stats['warnings']:
                            st.divider()
                            with st.expander(
                                f"⚠️ Peringatan ({len(stats['warnings'])})",
                                expanded=False
                            ):
                                for idx, warning in enumerate(stats['warnings'], 1):
                                    st.warning(f"{idx}. {warning}", icon="⚠️")
                        
                        if stats['errors']:
                            st.divider()
                            with st.expander(
                                f"❌ Error ({len(stats['errors'])})",
                                expanded=True
                            ):
                                for idx, error in enumerate(stats['errors'], 1):
                                    st.error(f"{idx}. {error}", icon="❌")
                        
                        # ============ PREVIEW XML ============
                        st.divider()
                        
                        with st.expander("📄 Preview Hasil XML", expanded=False):
                            lines = xml_result.split('\n')
                            total_lines = len(lines)
                            
                            if total_lines > PREVIEW_LINES:
                                preview = '\n'.join(lines[:PREVIEW_LINES])
                                st.code(preview, language="xml")
                                st.info(
                                    f"📊 Preview menampilkan {PREVIEW_LINES} dari {total_lines} baris\n\n"
                                    f"Silakan unduh file untuk melihat keseluruhan"
                                )
                            else:
                                st.code(xml_result, language="xml")
                            
                            # Show raw stats
                            st.markdown("---")
                            st.markdown("**📈 Detail XML:**")
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Baris", total_lines)
                            with col2:
                                st.metric("Ukuran", f"{len(xml_result) / 1024:.2f} KB")
                            with col3:
                                st.metric("Karakter", len(xml_result))
                
                except ConversionError as e:
                    st.error(
                        f"❌ **Conversion Error:**\n\n"
                        f"{str(e)}\n\n"
                        f"Mohon periksa format file DOCX Anda"
                    )
                    logger.error(f"Conversion error: {str(e)}")
                
                except Exception as e:
                    st.error(
                        f"❌ **Kesalahan Sistem:**\n\n"
                        f"{str(e)}\n\n"
                        f"Silakan hubungi administrator"
                    )
                    logger.error(f"Unexpected error: {str(e)}", exc_info=True)

# ============ FOOTER ============
st.divider()
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**ℹ️ Tentang**")
    st.markdown("Docx to Moodle XML Converter v2.0\n\nConversi otomatis file Word ke format Moodle")

with col2:
    st.markdown("**🔧 Teknologi**")
    st.markdown("• Python\n• Streamlit\n• docx2python\n• PIL")

with col3:
    st.markdown("**📝 License**")
    st.markdown("MIT License - Open Source")
