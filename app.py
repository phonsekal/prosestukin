import streamlit as st
import pandas as pd
import numpy as np
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin Otomatis", layout="wide")

st.title("📂 Pengolah Data Master Pembayaran")
st.write("Unggah semua file sekaligus (Master & File Tukin). Sistem akan otomatis mengenali file Master.")

# --- FUNGSI PEMROSESAN FILE SUMBER (TUKIN) ---
def process_source_file(file):
    try:
        # 1. Baca mentah untuk cari baris header (NIP)
        df_raw = pd.read_excel(file, header=None)
        header_row_idx = None
        for i, row in df_raw.iterrows():
            if "NIP" in row.values:
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return None
        
        # 2. Baca ulang dengan header yang benar
        df = pd.read_excel(file, skiprows=header_row_idx)
        
        # 3. Paksa nama kolom unik (Mencegah TypeError: arg must be a list...)
        # Jika ada dua kolom 'Potongan', akan menjadi 'Potongan', 'Potongan.1'
        new_cols = []
        counts = {}
        for col in df.columns:
            clean_col = str(col).replace('\n', ' ').strip()
            if clean_col in counts:
                counts[clean_col] += 1
                new_cols.append(f"{clean_col}.{counts[clean_col]}")
            else:
                counts[clean_col] = 0
                new_cols.append(clean_col)
        df.columns = new_cols

        # 4. Cari Kolom Penting
        nip_col = [c for c in df.columns if 'NIP' in c][0]
        tunjangan_cols = [c for c in df.columns if 'Tunjangan' in c]
        bruto_col = tunjangan_cols[0] if tunjangan_cols else None
        
        # Ambil semua kolom yang mengandung kata 'Potongan'
        potongan_cols = [c for c in df.columns if 'Potongan' in c and 'Total' not in c]
        
        if not bruto_col:
            return None

        # 5. Konversi ke Numerik & Kalkulasi
        # Kita gunakan .iloc untuk memastikan kita memproses Series tunggal
        df[bruto_col] = pd.to_numeric(df[bruto_col], errors='coerce').fillna(0)
        
        # Jumlahkan potongan secara manual per baris
        total_potongan_series = pd.Series(0, index=df.index)
        for col in potongan_cols:
            total_potongan_series += pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        df['Final_Potongan'] = total_potongan_series
        
        # 6. Bersihkan NIP (Hapus .0)
        df[nip_col] = df[nip_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        # Ambil hasil
        res = df[[nip_col, bruto_col, 'Final_Potongan']].copy()
        res.columns = ['NIP', 'Bruto_Src', 'Potongan_Src']
        
        return res
    except Exception as e:
        st.error(f"Gagal memproses {file.name}: {e}")
        return None

# --- ANTARMUKA UNGGAH TUNGGAL ---
uploaded_files = st.file_uploader(
    "Unggah semua file Excel di sini (Termasuk MASTER PEMBAYARAN.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    master_data = None
    source_datasets = []
    
    # Pisahkan Master dan Source berdasarkan nama file
    for f in uploaded_files:
        if "MASTER PEMBAYARAN" in f.name.upper():
            master_data = f
        else:
            processed = process_source_file(f)
            if processed is not None:
                source_datasets.append(processed)
                st.info(f"✅ File sumber terbaca: {f.name}")
            else:
                st.warning(f"⚠️ File diabaikan (NIP/Tunjangan tidak ketemu): {f.name}")

    if st.button("🚀 Proses & Gabungkan Data"):
        if master_data is None:
            st.error("File 'MASTER PEMBAYARAN.xlsx' tidak ditemukan dalam daftar unggahan!")
        elif not source_datasets:
            st.error("Tidak ada file sumber (Tukin) yang valid untuk diproses.")
        else:
            try:
                # 1. Baca Master
                df_master = pd.read_excel(master_data)
                # Normalisasi NIP Master
                nip_master_col = [c for c in df_master.columns if 'NIP' in c]
                if not nip_master_col:
                    st.error("Kolom 'NIP' tidak ada di file Master!")
                    st.stop()
                
                target_nip = nip_master_col[0]
                df_master[target_nip] = df_master[target_nip].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                
                # 2. Gabungkan semua sumber
                df_all_sources = pd.concat(source_datasets).drop_duplicates(subset=['NIP'], keep='first')
                
                # 3. Merge (VLOOKUP ala Python)
                df_final = pd.merge(df_master, df_all_sources, left_on=target_nip, right_on='NIP', how='left')
                
                # 4. Isi kolom yang diminta
                # Gunakan .fillna(0) agar tidak ada error saat pengurangan
                df_final['Nilai Bruto'] = df_final['Bruto_Src'].fillna(0)
                df_final['Nilai Potongan'] = df_final['Potongan_Src'].fillna(0)
                df_final['Nilai Bersih'] = df_final['Nilai Bruto'] - df_final['Nilai Potongan']
                
                # Hapus kolom sementara
                cols_to_drop = ['NIP', 'Bruto_Src', 'Potongan_Src']
                df_final = df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns])
                
                st.success("Berhasil! Data telah diperbarui.")
                st.dataframe(df_final.head(10))
                
                # 5. Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Update_Tukin')
                
                st.download_button(
                    label="📥 Unduh Hasil Akhir",
                    data=output.getvalue(),
                    file_name="HASIL_MASTER_PEMBAYARAN_UPDATED.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Terjadi kesalahan saat penggabungan: {e}")
