import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Pelindo PDF Tools", layout="wide")

st.title("⚓ Pelindo Manning Deployment Tools")

# Membuat dua tab untuk dua fungsi berbeda
tab1, tab2 = st.tabs(["📄 PDF to Excel (Layout Asli)", "📊 Rekap 3 Kolom (Final)"])

# --- TAB 1: CONVERT PDF KE EXCEL (LAYOUT ASLI) ---
with tab1:
    st.header("Konversi PDF ke Excel")
    st.info("Gunakan ini untuk mengubah file PDF asli menjadi Excel tanpa mengubah posisi kolom.")
    
    files_pdf = st.file_uploader("Unggah PDF Asli", type="pdf", accept_multiple_files=True, key="pdf_converter")
    
    if files_pdf:
        all_dfs = []
        for file in files_pdf:
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
                    if table:
                        df_page = pd.DataFrame(table).applymap(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)
                        all_dfs.append(df_page)
        
        if all_dfs:
            final_asli = pd.concat(all_dfs, ignore_index=True)
            st.dataframe(final_asli)
            
            output_asli = BytesIO()
            with pd.ExcelWriter(output_asli, engine='xlsxwriter') as writer:
                final_asli.to_excel(writer, index=False, header=False)
            
            st.download_button("📥 Download Excel Layout Asli", output_asli.getvalue(), "layout_asli.xlsx")

# --- TAB 2: REKAP 3 KOLOM (DARI CSV/EXCEL HASIL CONVERT) ---
with tab2:
    st.header("Rekap 3 Kolom")
    st.info("Unggah file Excel/CSV hasil dari Tab 1 untuk diekstrak menjadi 3 kolom: ITV, ID, dan Nama.")
    
    file_recap = st.file_uploader("Unggah File Excel/CSV Hasil Convert", type=["csv", "xlsx"], key="recap_tool")
    
    def extract_3_columns(df):
        extracted = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                cell_val = str(df.iloc[i, j]).strip()
                
                # Cari nomor ITV (3 digit angka)
                if re.fullmatch(r'\d{3}', cell_val):
                    itv_no = cell_val
                    
                    # Cari Operator di baris yang sama atau 1 baris di bawahnya
                    found = False
                    rows_to_check = [i, i+1] if i+1 < len(df) else [i]
                    
                    for row_idx in rows_to_check:
                        for col_idx in range(len(df.columns)):
                            target_val = str(df.iloc[row_idx, col_idx]).strip()
                            # Cari pola ID 4 digit + Nama
                            match = re.search(r'(\d{4})\s+([A-Z\s\.,]+)', target_val)
                            if match:
                                extracted.append({
                                    "Nomor ITV": itv_no,
                                    "Nomor Operator": match.group(1),
                                    "Nama Operator": match.group(2).strip()
                                })
                                found = True; break
                        if found: break
        return pd.DataFrame(extracted).drop_duplicates()

    if file_recap:
        if file_recap.name.endswith('.csv'):
            df_to_process = pd.read_csv(file_recap, header=None)
        else:
            df_to_process = pd.read_excel(file_recap, header=None)
            
        df_final = extract_3_columns(df_to_process)
        
        if not df_final.empty:
            st.success(f"Ditemukan {len(df_final)} data operator.")
            st.dataframe(df_final, use_container_width=True)
            
            output_3col = BytesIO()
            with pd.ExcelWriter(output_3col, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Rekap_Final')
            
            st.download_button("📥 Download Rekap 3 Kolom", output_3col.getvalue(), "rekap_final_3kolom.xlsx")
        else:
            st.warning("Data tidak ditemukan. Pastikan file yang diupload adalah hasil konversi dari Tab 1.")2
