import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Gabugn File Tukin", layout="wide")

st.title("📂 Gabung File Tukin")
st.write("Proses penggabungan data tunjangan kinerja pegawai")

# --- FUNGSI MEMBERSIHKAN ANGKA ---
def clean_numeric(value):
    if pd.isna(value): return 0.0
    # Jika sudah angka, langsung kembalikan
    if isinstance(value, (int, float)): return float(value)
    # Jika string, hapus karakter selain angka, koma, dan titik
    s = str(value).replace('Rp', '').replace(' ', '').replace('.', '').replace(',', '.')
    try:
        return float(re.sub(r'[^0-9.]', '', s))
    except:
        return 0.0

# --- FUNGSI PEMROSESAN FILE SUMBER ---
def process_source_file(file):
    try:
        # Baca semua sheet (default sheet pertama) sebagai string
        df_raw = pd.read_excel(file, header=None, dtype=str)
        
        # 1. Temukan Baris Header dan Kolom NIP
        header_row_idx = None
        nip_col_idx = None
        
        for i, row in df_raw.iterrows():
            row_list = [str(x).upper() for x in row.values]
            if "NIP" in row_list:
                header_row_idx = i
                nip_col_idx = row_list.index("NIP")
                break
        
        if header_row_idx is None:
            st.error(f"❌ Kolom 'NIP' tidak ditemukan di file: {file.name}")
            return None
        
        # 2. Ambil Data mulai dari baris header
        df = pd.read_excel(file, skiprows=header_row_idx, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 3. Cari Indeks Bruto (Tunjangan Pertama)
        bruto_idx = None
        for i, col in enumerate(df.columns):
            if "TUNJANGAN" in col.upper():
                bruto_idx = i
                break
        
        if bruto_idx is None:
            st.error(f"❌ Kolom 'Tunjangan' tidak ditemukan di file: {file.name}")
            return None

        # 4. Cari Indeks Bersih (Target: Kolom ke-6 atau kata kunci 'Diterima'/'Bersih')
        # Cek berdasarkan urutan (Bruto=1, Bersih=6 -> idx + 5)
        bersih_idx = bruto_idx + 5
        
        # Verifikasi apakah kolom tersebut benar (biasanya berisi kata 'Diterima' atau 'Bersih')
        # Jika tidak, kita cari kolom terdekat yang punya kata kunci tersebut
        if bersih_idx >= len(df.columns) or "DITERIMA" not in str(df.columns[bersih_idx]).upper():
            for i, col in enumerate(df.columns):
                if "DITERIMA" in col.upper() or "BERSIH" in col.upper():
                    bersih_idx = i
                    break

        # 5. Ekstrak dan Bersihkan Data
        results = []
        for _, row in df.iterrows():
            nip = str(row.iloc[nip_col_idx]).replace('.0', '').strip()
            # Abaikan jika NIP kosong atau bukan angka panjang
            if not nip or nip == 'nan' or len(nip) < 5:
                continue
                
            raw_bruto = row.iloc[bruto_idx]
            raw_bersih = row.iloc[bersih_idx]
            
            val_bruto = clean_numeric(raw_bruto)
            val_bersih = clean_numeric(raw_bersih)
            val_potongan = val_bruto - val_bersih
            
            results.append({
                'KEY_NIP': nip,
                'VAL_BRUTO': val_bruto,
                'VAL_BERSIH': val_bersih,
                'VAL_POTONGAN': val_potongan
            })
            
        return pd.DataFrame(results)

    except Exception as e:
        st.error(f"⚠️ Gagal memproses {file.name}: {str(e)}")
        return None

# --- UI ---
uploaded_files = st.file_uploader(
    "Unggah File Master & Semua File Tukin", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    master_file_obj = None
    source_datasets = []

    for f in uploaded_files:
        if "MASTER PEMBAYARAN" in f.name.upper():
            master_file_obj = f
        else:
            processed = process_source_file(f)
            if processed is not None and not processed.empty:
                source_datasets.append(processed)
                st.info(f"✅ Berhasil Ekstrak {len(processed)} data dari: {f.name}")

    if st.button("🚀 Jalankan Sinkronisasi"):
        if master_file_obj is None:
            st.error("File 'MASTER PEMBAYARAN.xlsx' belum diunggah!")
        elif not source_datasets:
            st.error("Tidak ada data yang bisa diekstrak dari file sumber.")
        else:
            try:
                # 1. Baca Master
                df_master = pd.read_excel(master_file_obj, dtype=str)
                
                # Cari kolom NIP di master
                m_nip_cols = [c for c in df_master.columns if 'NIP' in c.upper()]
                if not m_nip_cols:
                    st.error("Kolom NIP tidak ditemukan di file Master!")
                    st.stop()
                
                m_nip_name = m_nip_cols[0]
                df_master['MATCH_KEY'] = df_master[m_nip_name].astype(str).str.replace('.0', '', regex=False).str.strip()

                # 2. Gabungkan Semua Sumber
                df_all_sources = pd.concat(source_datasets).drop_duplicates(subset=['KEY_NIP'], keep='first')

                # 3. Merge
                df_final = pd.merge(
                    df_master, 
                    df_all_sources, 
                    left_on='MATCH_KEY', 
                    right_on='KEY_NIP', 
                    how='left'
                )

                # 4. Update Kolom
                df_final['Nilai Bruto'] = df_final['VAL_BRUTO'].fillna(0)
                df_final['Nilai Bersih'] = df_final['VAL_BERSIH'].fillna(0)
                df_final['Nilai Potongan'] = df_final['VAL_POTONGAN'].fillna(0)

                # 5. Bersihkan kolom temporary
                to_drop = ['MATCH_KEY', 'KEY_NIP', 'VAL_BRUTO', 'VAL_BERSIH', 'VAL_POTONGAN']
                df_final = df_final.drop(columns=[c for c in to_drop if c in df_final.columns])

                st.success("Sinkronisasi Selesai!")
                st.dataframe(df_final.head(10))

                # 6. Export dengan Proteksi NIP
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Update')
                    workbook  = writer.book
                    worksheet = writer.sheets['Update']
                    text_format = workbook.add_format({'num_format': '@'})
                    
                    for i, col in enumerate(df_final.columns):
                        if 'NIP' in col.upper():
                            worksheet.set_column(i, i, None, text_format)

                st.download_button(
                    label="📥 Unduh Hasil",
                    data=output.getvalue(),
                    file_name="MASTER_PEMBAYARAN_FINAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error Final: {e}")
