import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Konfigurasi Halaman
st.set_page_config(page_title="PDF to Excel Converter - Pelindo", layout="wide")

st.title("🚢 Pelindo PDF to Excel Converter")
st.markdown("Unggah file PDF Manning & Deployment untuk rekap otomatis tanpa merusak tata letak kolom.")

# Upload File
uploaded_files = st.file_uploader("Pilih satu atau beberapa file PDF", type="pdf", accept_multiple_files=True)

def extract_table_from_pdf(file):
    all_rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            # Menggunakan strategi 'lines' untuk menjaga presisi kolom seperti di PDF asli
            table = page.extract_table({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_y_tolerance": 5,
                "intersection_x_tolerance": 5,
            })
            if table:
                all_rows.extend(table)
    return all_rows

if uploaded_files:
    combined_data = []
    
    for file in uploaded_files:
        data = extract_table_from_pdf(file)
        if data:
            # Membersihkan teks dari karakter enter (\n) agar rapi di Excel
            df = pd.DataFrame(data)
            df = df.applymap(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)
            combined_data.append(df)
    
    if combined_data:
        # Menggabungkan semua PDF menjadi satu sheet
        final_df = pd.concat(combined_data, ignore_index=True)
        
        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file!")
        st.dataframe(final_df) # Preview data

        # Proses Download ke Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, header=False, sheet_name='Rekap_ITV')
        
        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=output.getvalue(),
            file_name="rekap_pelindo_itv.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
