import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="PDF to Excel Converter (ITV)", layout="wide")

st.title("⚓ Rekap Manning ITV ke Excel")
st.write("Unggah file PDF Manning & Deployment untuk dikonversi ke format Excel.")

uploaded_files = st.file_uploader("Pilih file PDF", type="pdf", accept_multiple_files=True)

def process_pdf(file):
    all_data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            # Ekstrak tabel dengan konfigurasi agar teks tidak berantakan
            table = page.extract_table({
                "vertical_strategy": "lines", 
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 10
            })
            
            if table:
                df = pd.DataFrame(table)
                # Membersihkan karakter '\n' yang sering muncul di PDF
                df = df.applymap(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)
                all_data.append(df)
    
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return None

if uploaded_files:
    # Gabungkan semua PDF yang diunggah jika ada banyak
    final_df_list = []
    for uploaded_file in uploaded_files:
        df_result = process_pdf(uploaded_file)
        if df_result is not None:
            final_df_list.append(df_result)
    
    if final_df_list:
        combined_df = pd.concat(final_df_list, ignore_index=True)
        
        st.success(f"Berhasil memproses {len(uploaded_files)} file!")
        st.dataframe(combined_df) # Tampilkan preview

        # Tombol Download Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, header=False, sheet_name='Rekap_Manning')
        
        st.download_button(
            label="Download Excel (.xlsx)",
            data=output.getvalue(),
            file_name="rekap_manning_deployment.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
