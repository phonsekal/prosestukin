import streamlit as st
import pandas as pd
import numpy as np
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Update Master Pembayaran", layout="wide")

st.title("📂 Pengolah Data Master Pembayaran")
st.write("Gunakan aplikasi ini untuk sinkronisasi data Tukin ke Master Pembayaran.")

# --- FUNGSI PEMROSESAN FILE SUMBER ---
def process_source_file(file):
    try:
        # Baca dengan dtype=str untuk mencegah pembulatan NIP (mencegah 000 di belakang)
        df_raw = pd.read_excel(file, header=None, dtype=str)
        
        header_row_idx = None
        for i, row in df_raw.iterrows():
            if "NIP" in row.values:
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return None
        
        # Baca ulang dengan header yang benar, tetap paksa dtype=str
        df = pd.read_excel(file, skiprows=header_row_idx, dtype=str)
        
        # Buat nama kolom unik
        new_cols = []
        counts = {}
        for col in df.columns:
            c = str(col).replace('\n', ' ').strip()
            if c in counts:
                counts[c] += 1
                new_cols.append(f"{c}_dup_{counts[c]}")
            else:
                counts[c] = 0
                new_cols.append(c)
        df.columns = new_cols

        # Identifikasi Kolom
        nip_col = [c for c in df.columns if 'NIP' in c][0]
        tunjangan_cols = [c for c in df.columns if 'Tunjangan' in c]
        bruto_col = tunjangan_cols[0] if tunjangan_cols else None
        potongan_cols = [c for c in df.columns if 'Potongan' in c and 'Total' not in c]

        if not bruto_col:
            return None

        # Konversi Nilai ke Numeric secara aman
        def safe_to_numeric(series):
            if isinstance(series, pd.DataFrame):
                series = series.iloc[:, 0]
            return pd.to_numeric(series, errors='coerce').fillna(0)

        bruto_values = safe_to_numeric(df[bruto_col])
        total_potongan_values = np.zeros(len(df))
        for p_col in potongan_cols:
            total_potongan_values += safe_to_numeric(df[p_col]).values
        
        # Bersihkan NIP (Hapus .0 dan spasi)
        source_nip = df[nip_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        res = pd.DataFrame({
            'KEY_NIP': source_nip,
            'VAL_BRUTO': bruto_values,
            'VAL_POTONGAN': total_potongan_values
        })
        
        # Hapus baris yang NIP-nya kosong
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
                # 1. Baca Master - PAKSA NIP JADI STRING
                df_master = pd.read_excel(master_file_obj, dtype=str)
                
                # Cari kolom NIP di master
                master_nip_col_name = [c for c in df_master.columns if 'NIP' in c][0]
                
                # Buat kunci pencocokan yang bersih
                df_master['KUNCI_MATCH'] = df_master[master_nip_col_name].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                # 2. Gabungkan Semua Sumber (Drop duplicates agar tidak double count)
                df_all_sources = pd.concat(source_datasets).drop_duplicates(subset=['KEY_NIP'], keep='first')

                # 3. Merge (VLOOKUP)
                df_final = pd.merge(
                    df_master, 
                    df_all_sources, 
                    left_on='KUNCI_MATCH', 
                    right_on='KEY_NIP', 
                    how='left'
                )

                # 4. Update Nilai (Gunakan fillna(0) agar tidak NaN)
                # Kita pastikan kolom target terisi nilai dari file sumber
                df_final['Nilai Bruto'] = df_final['VAL_BRUTO'].fillna(0)
                df_final['Nilai Potongan'] = df_final['VAL_POTONGAN'].fillna(0)
                
                # Pastikan tipe data numeric sebelum pengurangan
                df_final['Nilai Bruto'] = pd.to_numeric(df_final['Nilai Bruto'], errors='coerce').fillna(0)
                df_final['Nilai Potongan'] = pd.to_numeric(df_final['Nilai Potongan'], errors='coerce').fillna(0)
                df_final['Nilai Bersih'] = df_final['Nilai Bruto'] - df_final['Nilai Potongan']

                # 5. BERSIHKAN Kolom Tambahan
                cols_to_drop = ['KUNCI_MATCH', 'KEY_NIP', 'VAL_BRUTO', 'VAL_POTONGAN']
                df_final = df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns])

                st.success("Berhasil Sinkronisasi!")
                st.dataframe(df_final.head(10))

                # 6. EXPORT DENGAN ENGINE XLSXWRITER UNTUK MEMPERTAHANKAN STRING
                output = io.BytesIO()
                # Gunakan engine xlsxwriter
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Update_Tukin')
                    
                    # Akses objek xlsxwriter untuk memformat kolom NIP sebagai teks
                    workbook  = writer.book
                    worksheet = writer.sheets['Update_Tukin']
                    
                    # Format teks agar NIP tidak berubah jadi 000
                    text_format = workbook.add_format({'num_format': '@'})
                    
                    # Cari indeks kolom NIP untuk diformat
                    for i, col in enumerate(df_final.columns):
                        if 'NIP' in col.upper():
                            # Format kolom tersebut (baris 1 sampai akhir)
                            worksheet.set_column(i, i, None, text_format)

                st.download_button(
                    label="📥 Unduh Hasil Akhir (Presisi NIP Terjaga)",
                    data=output.getvalue(),
                    file_name="MASTER_PEMBAYARAN_TERUPDATE.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Terjadi error: {e}")
