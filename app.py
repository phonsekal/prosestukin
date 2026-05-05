import streamlit as st
import pandas as pd
import numpy as np
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Update Master Pembayaran", layout="wide")

st.title("📂 Pengolah Data Master Pembayaran")
st.write("Unggah file Master dan file-file sumber untuk mengisi Nilai Bruto, Potongan, dan Bersih secara otomatis.")

# --- FUNGSI PEMROSESAN FILE SUMBER ---
def process_source_file(file):
    try:
        # Baca file tanpa header dulu untuk mencari letak 'NIP'
        df_raw = pd.read_excel(file, header=None)
        
        header_row_idx = 0
        found_nip = False
        for i, row in df_raw.iterrows():
            if "NIP" in row.values:
                header_row_idx = i
                found_nip = True
                break
        
        if not found_nip:
            st.warning(f"Kolom 'NIP' tidak ditemukan di file: {file.name}")
            return None
        
        # Baca ulang dengan baris header yang benar
        df = pd.read_excel(file, skiprows=header_row_idx)
        
        # Penanganan Nama Kolom Duplikat (Penyebab TypeError)
        cols = []
        count = {}
        for col in df.columns:
            c = str(col).replace('\n', ' ').strip()
            if c in count:
                count[c] += 1
                cols.append(f"{c}.{count[c]}")
            else:
                count[c] = 0
                cols.append(c)
        df.columns = cols

        # Identifikasi Kolom NIP
        nip_col = [c for c in df.columns if 'NIP' in c][0]
        
        # Identifikasi Kolom Tunjangan (Bruto)
        tunjangan_cols = [c for c in df.columns if 'Tunjangan' in c]
        bruto_col = tunjangan_cols[0] if tunjangan_cols else None
        
        # Identifikasi Kolom Potongan
        # Mengambil semua kolom yang ada kata 'Potongan' tapi bukan 'Total Potongan'
        potongan_cols = [c for c in df.columns if 'Potongan' in c and 'Total' not in c]
        
        if not bruto_col:
            st.warning(f"Kolom 'Tunjangan' tidak ditemukan di file: {file.name}")
            return None

        # Konversi ke numerik dengan aman (Series per Series)
        df[bruto_col] = pd.to_numeric(df[bruto_col], errors='coerce').fillna(0)
        
        for col in potongan_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Hitung Total Potongan per baris
        df['Sum_Potongan'] = df[potongan_cols].sum(axis=1)
        
        # Seleksi data akhir
        res = df[[nip_col, bruto_col, 'Sum_Potongan']].copy()
        res.columns = ['NIP', 'Bruto_Src', 'Potongan_Src']
        
        # Normalisasi NIP (Hapus .0 jika terbaca sebagai float)
        res['NIP'] = res['NIP'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        return res
    except Exception as e:
        st.error(f"Error saat memproses {file.name}: {e}")
        return None

# --- ANTARMUKA PENGGUNA (UI) ---
col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader("1. Unggah MASTER PEMBAYARAN.xlsx", type=["xlsx"])

with col2:
    source_files = st.file_uploader("2. Unggah File-File Sumber (Tukin/Rekap)", type=["xlsx"], accept_multiple_files=True)

if master_file and source_files:
    if st.button("🚀 Jalankan Proses"):
        # 1. Baca Master
        try:
            df_master = pd.read_excel(master_file)
            # Normalisasi NIP Master
            if 'NIP' in df_master.columns:
                df_master['NIP'] = df_master['NIP'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            else:
                st.error("Kolom 'NIP' tidak ditemukan di file Master!")
                st.stop()
        except Exception as e:
            st.error(f"Gagal membaca file Master: {e}")
            st.stop()
        
        # 2. Proses semua file sumber
        all_sources = []
        for f in source_files:
            processed = process_source_file(f)
            if processed is not None:
                all_sources.append(processed)
        
        if all_sources:
            # Gabungkan semua data sumber menjadi satu referensi
            df_all_sources = pd.concat(all_sources).drop_duplicates(subset=['NIP'], keep='first')
            
            # 3. Gabungkan ke Master (Left Join)
            df_final = pd.merge(df_master, df_all_sources, on='NIP', how='left')
            
            # 4. Update Kolom Target
            df_final['Nilai Bruto'] = df_final['Bruto_Src'].fillna(0)
            df_final['Nilai Potongan'] = df_final['Potongan_Src'].fillna(0)
            df_final['Nilai Bersih'] = df_final['Nilai Bruto'] - df_final['Nilai Potongan']
            
            # Hapus kolom pembantu hasil merge
            df_final = df_final.drop(columns=['Bruto_Src', 'Potongan_Src'])
            
            # Tampilkan Hasil
            st.success("Pemrosesan Selesai!")
            st.subheader("Preview Hasil (5 Baris Pertama)")
            st.dataframe(df_final.head())
            
            # 5. Tombol Unduh
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Hasil_Update')
            
            st.download_button(
                label="📥 Unduh MASTER PEMBAYARAN Terupdate",
                data=output.getvalue(),
                file_name="MASTER_PEMBAYARAN_UPDATED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Tidak ada data yang berhasil diekstrak dari file sumber.")
else:
    st.info("Silakan unggah file Master dan minimal satu file sumber untuk memulai.")
