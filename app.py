import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Master Pembayaran Tukin", layout="wide", page_icon="📑")

st.title("📑 Generator Master Pembayaran Tukin")
st.markdown("Aplikasi ini menggabungkan data Tukin menjadi satu file **MASTER PEMBAYARAN.xlsx**.")

# Sidebar
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

def smart_numeric_cleaner(value):
    """Memastikan angka murni tidak rusak dan teks mata uang dibersihkan."""
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).replace('Rp', '').replace(' ', '')
        if ',' in text and '.' in text:
            text = text.replace('.', '').replace(',', '.')
        elif ',' in text:
            text = text.replace(',', '.')
        clean_str = re.sub(r'[^\d.]', '', text)
        return float(clean_str) if clean_str else 0.0
    except:
        return 0.0

uploaded_files = st.file_uploader("Upload semua file Tukin sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Proses & Buat Master Pembayaran"):
        all_data = []
        
        for file in uploaded_files:
            try:
                # 1. Cari baris header yang berisi 'NIP'
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                found_header = False
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        found_header = True
                        break
                
                if not found_header:
                    st.error(f"❌ Baris NIP tidak ditemukan di file: {file.name}")
                    continue

                # 2. Baca ulang file dengan header yang benar[cite: 1, 2]
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Tangani Merge Cells pada Header
                df.columns = pd.Series(df.columns).ffill().str.strip()

                # 4. Cari kolom NIP dan Nama secara manual (menghindari error ambiguitas)
                col_nip, col_nama = None, None
                for c in df.columns:
                    if 'NIP' in str(c).upper(): col_nip = c
                    if 'NAMA' in str(c).upper() and col_nama is None: col_nama = c

                if col_nip is None or col_nama is None:
                    st.error(f"❌ Kolom identitas tidak ditemukan di: {file.name}")
                    continue

                # 5. Buat DataFrame hasil
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(r'\D', '', regex=True)
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # 6. Pembersihan Nilai Keuangan (Bruto, Potongan, Bersih)
                # Menggunakan list comprehension untuk menghindari error Series Truth Value
                for target, keywords in [('BRUTO', 'BRUTO'), ('POTONGAN', 'POTONGAN'), ('BERSIH', 'BERSIH')]:
                    matched_cols = [c for c in df.columns if keywords in str(c).upper()]
                    
                    if len(matched_cols) > 0:
                        # Jumlahkan semua kolom yang cocok (menangani banyak kolom potongan)
                        temp_sum = pd.Series(0.0, index=df.index)
                        for m_col in matched_cols:
                            temp_sum += df[m_col].apply(smart_numeric_cleaner)
                        df_clean[target] = temp_sum
                    else:
                        df_clean[target] = 0.0

                # 7. Filter Baris Valid (NIP >= 9 digit)[cite: 1]
                df_clean = df_clean[df_clean['NIP'].str.len() >= 9].copy()
                
                # Paksa tipe data numerik[cite: 2]
                for col in ['BRUTO', 'POTONGAN', 'BERSIH']:
                    df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0.0)
                
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil memproses: {file.name}")
                
            except Exception as e:
                st.error(f"❌ Gagal memproses {file.name}: {str(e)}")

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            st.divider()
            st.subheader("📊 Pratinjau Master Pembayaran")
            st.dataframe(final_df, use_container_width=True)
            
            # Statistik Akhir[cite: 2]
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Pegawai", f"{len(final_df)} orang")
            c2.metric("Total Bruto", f"Rp {final_df['BRUTO'].sum():,.2f}")
            c3.metric("Total Bersih", f"Rp {final_df['BERSIH'].sum():,.2f}")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='MASTER_PEMBAYARAN')
            
            st.divider()
            st.download_button(
                label="📥 Download MASTER PEMBAYARAN.xlsx",
                data=output.getvalue(),
                file_name="MASTER PEMBAYARAN.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
