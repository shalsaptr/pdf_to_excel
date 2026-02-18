import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Pelindo ITV Recap", layout="wide")

st.title("⚓ Rekap ITV, ID, dan Nama Operator")
st.write("Aplikasi ini didesain khusus untuk format tabel Manning Pelindo.")

uploaded_files = st.file_uploader("Unggah PDF (01, 16, atau 28 Jan)", type="pdf", accept_multiple_files=True)

def clean_text(text):
    if text:
        return re.sub(r'\s+', ' ', str(text)).strip()
    return ""

def extract_itv_data(file):
    extracted_rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            # Gunakan extract_table dengan seting toleransi lebih tinggi
            table = page.extract_table({
                "vertical_strategy": "text", 
                "horizontal_strategy": "text",
            })
            
            if not table:
                # Fallback ke pencarian baris manual jika tabel gagal terdeteksi
                table = page.extract_table()

            if table:
                for row in table:
                    for cell in row:
                        if not cell: continue
                        val = clean_text(cell)
                        
                        # Cari pola: Angka ITV (3 digit) dan di dalamnya ada ID (4 digit)
                        # Seringkali di PDF Pelindo, ITV dan Nama tergabung dalam satu sel atau sel berdekatan
                        itv_match = re.search(r'\b(\d{3})\b', val)
                        op_match = re.search(r'(\d{4})\s+([A-Z\s\.,]+)', val)
                        
                        if itv_match and op_match:
                            extracted_rows.append({
                                "Nomor ITV": itv_match.group(1),
                                "Nomor Operator": op_match.group(1),
                                "Nama Operator": op_match.group(2).strip()
                            })
                        # Kondisi jika ITV dan Operator terpisah kolom (ITV di sel ini, Op di sel lain)
                        elif re.fullmatch(r'\d{3}', val):
                            itv_no = val
                            # Cari di sel lain dalam baris yang sama yang punya 4 digit ID
                            for other_cell in row:
                                other_val = clean_text(other_cell)
                                op_only_match = re.search(r'(\d{4})\s+([A-Z\s\.,]+)', other_val)
                                if op_only_match:
                                    extracted_rows.append({
                                        "Nomor ITV": itv_no,
                                        "Nomor Operator": op_only_match.group(1),
                                        "Nama Operator": op_only_match.group(2).strip()
                                    })
    return extracted_rows

if uploaded_files:
    all_results = []
    for f in uploaded_files:
        data = extract_itv_data(f)
        all_results.extend(data)
    
    if all_results:
        df = pd.DataFrame(all_results).drop_duplicates()
        st.success(f"Ditemukan {len(df)} data.")
        st.dataframe(df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data_ITV')
        
        st.download_button("📥 Download Excel (3 Kolom)", output.getvalue(), "rekap_itv_final.xlsx")
    else:
        st.warning("Data tidak ditemukan. Cobalah untuk memeriksa apakah file PDF tersebut bisa di-copy teksnya secara manual.")
