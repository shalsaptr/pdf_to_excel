import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="ITV Parser Pro", layout="wide")

st.title("⚓ ITV Operator Extractor")
st.write("Ekstraksi otomatis ITV, ID, dan Nama Operator.")

uploaded_files = st.file_uploader("Upload PDF", type="pdf", accept_multiple_files=True)

def extract_data_flexible(file):
    results = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            # Ambil semua kata beserta posisinya
            words = page.extract_words()
            
            for i, word in enumerate(words):
                text = word['text'].strip()
                
                # 1. Cari Nomor ITV (2-3 digit angka, misal 238, 255, 307)
                if re.match(r'^\d{3}$', text):
                    itv_no = text
                    
                    # 2. Cari Operator di kata-kata setelahnya (mencari pola ID 4 digit)
                    # Kita scan hingga 15 kata ke depan untuk menemukan ID Operator
                    for j in range(i + 1, min(i + 15, len(words))):
                        potential_op = words[j]['text'].strip()
                        
                        # Cek jika menemukan ID 4 digit (seperti 7230, 7111, dll)
                        if re.match(r'^\d{4}$', potential_op):
                            op_id = potential_op
                            
                            # Ambil nama (biasanya 2-4 kata setelah ID)
                            name_parts = []
                            for k in range(j + 1, min(j + 5, len(words))):
                                # Berhenti jika ketemu angka lagi (berarti sudah ITV baru/ID baru)
                                if words[k]['text'].isdigit():
                                    break
                                name_parts.append(words[k]['text'])
                            
                            op_name = " ".join(name_parts)
                            
                            results.append({
                                "Nomor ITV": itv_no,
                                "Nomor Operator": op_id,
                                "Nama Operator": op_name
                            })
                            break
    return results

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        records = extract_data_flexible(f)
        all_data.extend(records)
    
    if all_data:
        df = pd.DataFrame(all_data).drop_duplicates()
        st.success(f"Ditemukan {len(df)} data operator.")
        st.dataframe(df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Data_ITV')
        
        st.download_button("📥 Download Excel", output.getvalue(), "rekap_itv.xlsx")
    else:
        st.error("Data tidak ditemukan. Pastikan PDF memiliki teks yang dapat dibaca (bukan hasil scan gambar murni).")
