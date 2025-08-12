import streamlit as st
import google.generativeai as genai
import pandas as pd
from PIL import Image
from io import StringIO, BytesIO
import xlsxwriter

# =========================
# Konfigurasi Gemini API
# =========================
API_KEY = "AIzaSyDiZY999jWBbz0ddQKKNvS8v_Z2fGz6WaY"
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# Fungsi OCR dengan Gemini
def ocr_with_gemini(image):
    prompt = """
    Kamu adalah asisten OCR.
    Dari foto ini, ekstrak semua data barang dan total menjadi tabel CSV dengan format:
    Nama Barang,Total
    Pastikan tidak ada teks lain selain tabel CSV.
    Apabila ada satuan metrik seperti kg, dll. Hapus semua sehingga yang ditampilkan hanya angka.
    Semua angka bentuknya antara bilangan bulat/desimal. Jadi apabila ada jumlah seperti 1/2 atau yang lain tolong ubah jadi desimal.
    """
    response = model.generate_content([prompt, image])
    return response.text

# Fungsi untuk konversi CSV text jadi DataFrame
def csv_to_dataframe(csv_text):
    csv_io = StringIO(csv_text)
    df = pd.read_csv(csv_io)
    return df

# Fungsi membuat Excel terpisah menjadi 2 blok
def to_excel_split(df, rows_per_block=43):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    # Konversi DataFrame ke list (header + data)
    data = [list(df.columns)] + df.values.tolist()

    # Blok 1
    for row_idx, row_data in enumerate(data[:rows_per_block+1]):  # +1 untuk header
        for col_idx, value in enumerate(row_data):
            fmt = bold if row_idx == 0 else None
            worksheet.write(row_idx, col_idx, value, fmt)

    # Blok 2 mulai kolom D (index 3)
    start_col = 3
    for row_idx, row_data in enumerate(data[0:1] + data[rows_per_block+1:]):  
        for col_idx, value in enumerate(row_data):
            fmt = bold if row_idx == 0 else None
            worksheet.write(row_idx, start_col + col_idx, value, fmt)

    workbook.close()
    return output.getvalue()

# =========================
# UI Streamlit
# =========================
st.title("OCR Gemini - Tabel, Copy & Download Excel (Split 2 Blok)")
uploaded_file = st.file_uploader("Upload gambar", type=["jpg", "jpeg", "png"])

if uploaded_file:
    image = Image.open(uploaded_file)
    st.image(image, caption="Gambar yang diunggah", use_column_width=True)

    with st.spinner("Menganalisis gambar..."):
        ocr_result = ocr_with_gemini(image)

    st.subheader("Hasil OCR Mentah")
    st.text(ocr_result)

    try:
        df = csv_to_dataframe(ocr_result)
        st.subheader("Tabel Hasil")
        st.dataframe(df)

        # Persiapkan data
        total_1_44 = "\n".join(df['Total'].iloc[0:43].astype(str).tolist())
        total_45_end = "\n".join(df['Total'].iloc[43:].astype(str).tolist())

        # Tampilkan data sebagai code block jika button diklik
        # Alternative: Tampilkan langsung tanpa tombol (always visible)
        st.markdown("---")
        st.subheader("📋 Data Siap Copy (Always Visible)")
        
        col5, col6 = st.columns(2)
        
        with col5:
            st.write("**TOTAL 1–44:**")
            with st.container():
                st.code(total_1_44, language=None)

        with col6:
            st.write("**TOTAL 45–akhir:**")
            with st.container():
                st.code(total_45_end, language=None)

        # Backup: Download files
        st.markdown("---")
        st.subheader("💾 Download sebagai File Text (Backup)")
        col7, col8 = st.columns(2)

        # Excel download
        st.markdown("---")
        st.subheader("📊 Download Excel")
        excel_bytes = to_excel_split(df, rows_per_block=43)
        st.download_button(
            label="💾 Download Excel (2 Blok)",
            data=excel_bytes,
            file_name="hasil_ocr_split.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"❌ Gagal parsing hasil OCR menjadi tabel: {e}")