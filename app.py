import streamlit as st
import pandas as pd
import numpy as np
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Update Master Pembayaran", layout="wide")

st.title("📂 Pengolah Data Master Pembayaran")
st.write("Unggah semua file sekaligus. Sistem akan mendeteksi 'MASTER PEMBAYARAN.xlsx' secara otomatis.")

# --- FUNGSI PEMROSESAN FILE SUMBER ---
def process_source_file(file):
    try:
        # 1. Cari baris header (NIP)
        df_raw = pd.read_excel(file, header=None)
        header_row_idx = None
        for i, row in df_raw.iterrows():
            if "NIP" in row.values:
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return None
        
        # 2. Baca ulang dengan header benar
        df = pd.read_excel(file, skiprows=header_row_idx)
        
        # 3. Paksa semua nama kolom menjadi unik dan bersih
        # Menggunakan list comprehension untuk menangani duplikat sebelum kolom ditetapkan
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

        # 4. Identifikasi Nama Kolom
        nip_col = [c for c in df.columns if 'NIP' in c][0]
        tunjangan_cols = [c for c in df.columns if 'Tunjangan' in c]
        bruto_col = tunjangan_cols[0] if tunjangan_cols else None
        potongan_cols = [c for c in df.columns if 'Potongan' in c and 'Total' not in c]

        if not bruto_col:
            return None

        # 5. Konversi Numerik (MENGGUNAKAN CARA PALING AMAN)
        # Kita ambil series secara eksplisit untuk menghindari DataFrame duplikat
        def safe_to_numeric(series):
            # Jika series ternyata masih dataframe (akibat duplikasi pandas yang sangat bandel)
            if isinstance(series, pd.DataFrame):
                series = series.iloc[:, 0]
            return pd.to_numeric(series, errors='coerce').fillna(0)

        # Hitung Bruto
        bruto_values = safe_to_numeric(df[bruto_col])
        
        # Hitung Total Potongan
        total_potongan_values = np.zeros(len(df))
        for p_col in potongan_cols:
            total_potongan_values += safe_to_numeric(df[p_col]).values
        
        # 6. Bangun DataFrame Hasil
        res = pd.DataFrame({
            'NIP': df[nip_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip(),
            'Bruto_Src': bruto_values,
            'Potongan_Src': total_potongan_values
        })
        
        return res

    except Exception as e:
        st.error(f"⚠️ Gagal memproses {file.name}: {e}")
        return None

# --- UI UNGGAH ---
uploaded_files = st.file_uploader(
    "Unggah File Master & File Tukin (Bisa sekaligus banyak)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    master_file_obj = None
    source_datasets = []

    # Filter Master vs Sumber
    for f in uploaded_files:
        if "MASTER PEMBAYARAN" in f.name.upper():
            master_file_obj = f
        else:
            processed = process_source_file(f)
            if processed is not None:
                source_datasets.append(processed)
                st.info(f"✅ Berhasil membaca: {f.name}")

    if st.button("🚀 Jalankan Update Master"):
        if master_file_obj is None:
            st.error("File 'MASTER PEMBAYARAN.xlsx' tidak ditemukan!")
        elif not source_datasets:
            st.error("Tidak ada data valid dari file sumber untuk digabungkan.")
        else:
            try:
                # 1. Baca Master
                df_master = pd.read_excel(master_file_obj)
                
                # Cari kolom NIP di master
                master_nip_col = [c for c in df_master.columns if 'NIP' in c]
                if not master_nip_col:
                    st.error("Kolom 'NIP' tidak ditemukan di Master!")
                    st.stop()
                
                m_nip = master_nip_col[0]
                df_master[m_nip] = df_master[m_nip].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                # 2. Gabungkan Semua Sumber
                df_all_sources = pd.concat(source_datasets).drop_duplicates(subset=['NIP'], keep='first')

                # 3. Gabungkan (Merge)
                df_final = pd.merge(df_master, df_all_sources, left_on=m_nip, right_on='NIP', how='left')

                # 4. Update Kolom Target
                # Pastikan kolom target ada, jika tidak, akan dibuat otomatis
                df_final['Nilai Bruto'] = df_final['Bruto_Src'].fillna(0)
                df_final['Nilai Potongan'] = df_final['Potongan_Src'].fillna(0)
                df_final['Nilai Bersih'] = df_final['Nilai Bruto'] - df_final['Nilai Potongan']

                # Bersihkan kolom temporary
                cols_to_drop = ['NIP', 'Bruto_Src', 'Potongan_Src']
                df_final = df_final.drop(columns=[c for c in cols_to_drop if c in df_final.columns])

                st.success("Proses Berhasil!")
                st.dataframe(df_final.head(10))

                # 5. Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Update_Master')
                
                st.download_button(
                    label="📥 Unduh MASTER_PEMBAYARAN_UPDATED.xlsx",
                    data=output.getvalue(),
                    file_name="MASTER_PEMBAYARAN_UPDATED.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Terjadi error saat penggabungan: {e}")
