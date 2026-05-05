import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Update Master Pembayaran", layout="wide")

st.title("📂 Pengolah Data Master Pembayaran")
st.write("Unggah file Master dan file-file Tukin untuk mengisi Nilai Bruto, Potongan, dan Bersih secara otomatis.")

# 1. Upload Files
col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader("Unggah MASTER PEMBAYARAN.xlsx", type=["xlsx"])

with col2:
    source_files = st.file_uploader("Unggah File-File Sumber (Tukin/Rekap)", type=["xlsx"], accept_multiple_files=True)

def process_source_file(file):
    # Baca file, coba temukan header yang mengandung 'NIP'
    df_raw = pd.read_excel(file, header=None)
    
    # Cari baris mana yang mengandung kata 'NIP'
    header_row_idx = 0
    for i, row in df_raw.iterrows():
        if "NIP" in row.values:
            header_row_idx = i
            break
    
    # Baca ulang dengan header yang benar
    df = pd.read_excel(file, skiprows=header_row_idx)
    
    # Bersihkan nama kolom dari whitespace atau newline
    df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
    
    # Identifikasi Kolom
    # NIP
    nip_col = [c for c in df.columns if 'NIP' in c][0]
    
    # Tunjangan (Bruto) - ambil yang pertama muncul
    tunjangan_cols = [c for c in df.columns if 'Tunjangan' in c]
    bruto_col = tunjangan_cols[0] if tunjangan_cols else None
    
    # Potongan - semua yang ada kata 'Potongan' (kecuali jika ada 'Total Potongan' agar tidak double count)
    # Namun sesuai instruksi: "jumlahkan semua kolom yang ada kata potongan"
    potongan_cols = [c for c in df.columns if 'Potongan' in c and 'Total' not in c]
    
    if not bruto_col or not nip_col:
        return None

    # Konversi ke numerik
    df[bruto_col] = pd.to_numeric(df[bruto_col], errors='coerce').fillna(0)
    for col in potongan_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Hitung Total Potongan per baris
    df['Sum_Potongan'] = df[potongan_cols].sum(axis=1)
    
    # Ambil data relevan
    res = df[[nip_col, bruto_col, 'Sum_Potongan']].copy()
    res.columns = ['NIP', 'Bruto_Src', 'Potongan_Src']
    
    # Pastikan NIP string
    res['NIP'] = res['NIP'].astype(str).str.strip()
    
    return res

if master_file and source_files:
    if st.button("Proses Data"):
        # Baca Master
        df_master = pd.read_excel(master_file)
        # Pastikan NIP di master adalah string
        df_master['NIP'] = df_master['NIP'].astype(str).str.strip()
        
        # Gabungkan semua data dari source_files
        all_sources = []
        for f in source_files:
            processed = process_source_file(f)
            if processed is not None:
                all_sources.append(processed)
        
        if all_sources:
            df_all_sources = pd.concat(all_sources).drop_duplicates(subset=['NIP'], keep='first')
            
            # Gabungkan ke Master
            df_final = pd.merge(df_master, df_all_sources, on='NIP', how='left')
            
            # Isi kolom target
            df_final['Nilai Bruto'] = df_final['Bruto_Src'].fillna(0)
            df_final['Nilai Potongan'] = df_final['Potongan_Src'].fillna(0)
            df_final['Nilai Bersih'] = df_final['Nilai Bruto'] - df_final['Nilai Potongan']
            
            # Hapus kolom pembantu
            df_final = df_final.drop(columns=['Bruto_Src', 'Potongan_Src'])
            
            st.success("Berhasil memproses data!")
            st.dataframe(df_final.head())
            
            # Download Button
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Hasil Update')
            
            st.download_button(
                label="📥 Unduh MASTER PEMBAYARAN Terupdate",
                data=output.getvalue(),
                file_name="MASTER_PEMBAYARAN_UPDATED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Tidak ditemukan kolom Tunjangan atau NIP pada file sumber.")
