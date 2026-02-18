import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="ITV Recap Tool", layout="wide")

st.title("⚓ Rekap Final ITV & Operator")
st.write("Unggah file PDF atau CSV hasil ekstraksi sebelumnya untuk mendapatkan format 3 kolom.")

uploaded_file = st.file_uploader("Pilih file (PDF atau CSV)", type=["pdf", "csv"])

def extract_from_dataframe(df):
    final_data = []
    
    # Iterasi setiap kolom dan baris
    for col in df.columns:
        # Cari angka ITV (3 digit)
        for i, value in enumerate(df[col]):
            val_str = str(value).strip()
            
            # Pola ITV: 3 digit angka (seperti 238, 245, 252)
            if re.fullmatch(r'\d{3}', val_str):
                itv_no = val_str
                
                # Cari Operator di baris yang sama atau tepat di bawahnya
                # Kita cek sel di sekitar koordinat ITV ditemukan
                found_op = False
                
                # Cek baris yang sama, kolom berbeda
                for target_val in df.iloc[i]:
                    target_str = str(target_val).strip()
                    # Cari pola ID 4 digit + Nama (Contoh: 7230 HARI SETYO CAHYANTO)
                    match = re.search(r'(\d{4})\s+([A-Za-z\s\.,]+)', target_str)
                    if match:
                        final_data.append({
                            "Nomor ITV": itv_no,
                            "Nomor Operator": match.group(1),
                            "Nama Operator": match.group(2).strip()
                        })
                        found_op = True
                        break
                
                # Jika tidak ketemu di baris yang sama, cek baris di bawahnya
                if not found_op and i + 1 < len(df):
                    for target_val in df.iloc[i+1]:
                        target_str = str(target_val).strip()
                        match = re.search(r'(\d{4})\s+([A-Za-z\s\.,]+)', target_str)
                        if match:
                            final_data.append({
                                "Nomor ITV": itv_no,
                                "Nomor Operator": match.group(1),
                                "Nama Operator": match.group(2).strip()
                            })
                            break
                            
    return pd.DataFrame(final_data).drop_duplicates()

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df_input = pd.read_csv(uploaded_file, header=None)
    else:
        # Jika PDF, gunakan pdfplumber (pastikan requirements.txt sudah ada pdfplumber)
        import pdfplumber
        all_rows = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table: all_rows.extend(table)
        df_input = pd.DataFrame(all_rows)

    if not df_input.empty:
        df_result = extract_from_dataframe(df_input)
        
        st.success(f"Ditemukan {len(df_result)} baris data.")
        st.dataframe(df_result, use_container_width=True)

        # Download Button
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Rekap')
        
        st.download_button(
            label="📥 Download Excel 3 Kolom",
            data=output.getvalue(),
            file_name="rekap_itv_3kolom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
