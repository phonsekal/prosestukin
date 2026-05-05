import streamlit as st
import pandas as pd
import numpy as np
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Update Master Pembayaran", layout="wide")

st.title("📂 Pengolah Data Master Pembayaran")
st.write("Sinkronisasi Tukin berdasarkan posisi kolom (Bruto & Bersih).")

# --- FUNGSI PEMROSESAN FILE SUMBER ---
def process_source_file(file):
    try:
        # Baca dengan dtype=str untuk mencegah pembulatan NIP
        df_raw = pd.read_excel(file, header=None, dtype=str)
        
        header_row_idx = None
        for i, row in df_raw.iterrows():
            if "NIP" in row.values:
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return None
        
        # Baca ulang dengan header yang benar
        df = pd.read_excel(file, skiprows=header_row_idx, dtype=str)
        
        # Bersihkan nama kolom untuk identifikasi
        clean_cols = [str(c).replace('\n', ' ').strip() for c in df.columns]
        df.columns = clean_cols

        # 1. Identifikasi Indeks NIP
        nip_idx = next(i for i, c in enumerate(df.columns) if 'NIP' in c)
        
        # 2. Identifikasi Indeks Bruto (Kolom Tunjangan pertama)
        bruto_idx = next(i for i, c in enumerate(df.columns) if 'Tunjangan' in c)
        
        # 3. Identifikasi Indeks Bersih (Kolom ke-6 setelah Bruto)
        # Sesuai instruksi: Bruto dihitung 1, maka kolom ke-6 adalah (bruto_idx + 5)
        bersih_idx = bruto_idx + 5
        
        if bersih_idx >= len(df.columns):
            st.warning(f"File {file.name} tidak memiliki cukup kolom untuk mengambil Nilai Bersih.")
            return None

        # Helper untuk konversi angka
        def to_num(series):
            return pd.to_numeric(series, errors='coerce').fillna(0)

        # Ekstrak Nilai
        nip_series = df.iloc[:, nip_idx].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        val_bruto = to_num(df.iloc[:, bruto_idx])
        val_bersih = to_num(df.iloc[:, bersih_idx])
        val_potongan = val_bruto - val_bersih
        
        res = pd.DataFrame({
            'KEY_NIP': nip_series,
            'VAL_BRUTO': val_bruto,
            'VAL_BERSIH': val_bersih,
            'VAL_POTONGAN': val_potongan
        })
        
        # Hapus baris kosong
        res = res[res['KEY_NIP'] != 'nan']
        
        return res

    except Exception as e:
        st.error(f"⚠️ Gagal memproses {file.name}: {e}")
        return None

# --- UI ---
uploaded_files = st.file_uploader(
    "Unggah File Master & File Tukin sekaligus", 
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
            if processed is not None:
                source_datasets.append(processed)
                st.info(f"✅ Terbaca: {f.name}")

    if st.button("🚀 Jalankan Proses Update"):
        if master_file_obj is None:
            st.error("File 'MASTER PEMBAYARAN.xlsx' belum diunggah!")
        elif not source_datasets:
            st.error("Tidak ada data valid dari file sumber.")
        else:
            try:
                # 1. Baca Master
                df_master = pd.read_excel(master_file_obj, dtype=str)
                
                # Identifikasi kolom NIP asli di Master
                master_nip_col_name = [c for c in df_master.columns if 'NIP' in c][0]
                
                # Buat kunci pencocokan
                df_master['KUNCI_MATCH'] = df_master[master_nip_col_name].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                # 2. Gabungkan Semua Sumber
                df_all_sources = pd.concat(source_datasets).drop_duplicates(subset=['KEY_NIP'], keep='first')

                # 3. Merge
                df_final = pd.merge(
                    df_master, 
                    df_all_sources, 
                    left_on='KUNCI_MATCH', 
                    right_on='KEY_NIP', 
                    how='left'
                )

                # 4. Update Kolom Target
                # Mengisi kolom 'Nilai Bruto', 'Nilai Potongan', 'Nilai Bersih'
                df_final['Nilai Bruto'] = df_final['VAL_BRUTO'].fillna(0)
                df_final['Nilai Bersih'] = df_final['VAL_BERSIH'].fillna(0)
                df_final['Nilai Potongan'] = df_final['VAL_POTONGAN'].fillna(0)

                # 5. BERSIHKAN Kolom Tambahan
                cols_to_drop = ['KUNCI_MATCH', 'KEY_NIP', 'VAL_BRUTO', 'VAL_BERSIH', 'VAL_POTONGAN']
                df_final = df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns])

                st.success("Sinkronisasi Berhasil!")
                st.dataframe(df_final.head(10))

                # 6. EXPORT DENGAN PROTEKSI TEKS NIP
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Update_Tukin')
                    
                    workbook  = writer.book
                    worksheet = writer.sheets['Update_Tukin']
                    
                    # Format kolom NIP agar tetap sebagai teks (mencegah pembulatan ke 000)
                    text_format = workbook.add_format({'num_format': '@'})
                    
                    for i, col in enumerate(df_final.columns):
                        if 'NIP' in col.upper():
                            worksheet.set_column(i, i, None, text_format)

                st.download_button(
                    label="📥 Unduh MASTER_PEMBAYARAN_TERUPDATE.xlsx",
                    data=output.getvalue(),
                    file_name="MASTER_PEMBAYARAN_TERUPDATE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Terjadi error saat merge: {e}")
