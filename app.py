import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="ITV Recap Tool", layout="wide")

st.title("📊 Rekap Operator ITV (3 Kolom)")
st.write("Aplikasi ini mengekstrak data ITV, ID Operator, dan Nama Operator ke dalam format Excel bersih.")

uploaded_files = st.file_uploader("Unggah PDF Manning & Deployment", type="pdf", accept_multiple_files=True)

def clean_and_parse(text):
    if not text:
        return None, None
    # Membersihkan karakter baru dan spasi berlebih
    text = text.replace('\n', ' ').strip()
    # Mencari pola ID (4 digit angka) di awal teks
    match = re.match(r'^(\d{4})\s+(.*)', text)
    if match:
        return match.group(1), match.group(2)
    return None, text

def process_to_three_columns(file):
    extracted_data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines"
            })
            
            if table:
                for row in table:
                    # Iterasi setiap sel dalam baris untuk mencari pasangan ITV dan Operator
                    # Berdasarkan struktur PDF, biasanya ITV ada di baris atas dan Operator di bawahnya
                    # atau berada di kolom yang bersebelahan.
                    for i, cell in enumerate(row):
                        if cell and str(cell).isdigit() and len(str(cell)) <= 3:
                            itv_number = cell
                            # Cek sel di bawahnya atau di sebelahnya untuk data operator
                            # Kita cari pola 4 digit ID dalam tabel
                            for search_cell in row:
                                if search_cell and re.search(r'\d{4}', str(search_cell)):
                                    op_id, op_name = clean_and_parse(search_cell)
                                    if op_id:
                                        extracted_data.append({
                                            "Nomor ITV": itv_number,
                                            "Nomor Operator": op_id,
                                            "Nama Operator": op_name
                                        })
                                        break
    return extracted_data

if uploaded_files:
    all_records = []
    for f in uploaded_files:
        data = process_to_three_columns(f)
        all_records.extend(data)
    
    if all_records:
        df_final = pd.DataFrame(all_records).drop_duplicates()
        
        st.success(f"Berhasil mengekstrak {len(df_final)} baris data.")
        st.dataframe(df_final, use_container_width=True)

        # Tombol Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Rekap_ITV')
        
        st.download_button(
            label="📥 Download Excel 3 Kolom",
            data=output.getvalue(),
            file_name="rekap_itv_3_kolom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
